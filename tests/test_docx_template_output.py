# -*- coding: utf-8 -*-
"""P0-B regression: ``docx_template.py apply --to X -o Y`` must honour ``-o``.

The bug: the ``--to`` branch hard-coded ``target = Path(args.to)`` and rewrote it
in place, never reading ``args.output``. So the command in ``SKILL.md`` L80
(``apply tpl.json --to in.docx -o out.docx``) silently clobbered ``in.docx``,
created no ``out.docx``, left no ``.bak``, and still exited 0 -- the user only
found out later that the original document was gone.

What is asserted here is the measured post-fix behaviour, not a wish list:

* ``--to`` + ``-o``  -> writes ``-o``; the ``--to`` file keeps its exact bytes
                       (SHA256 compare -- size alone can collide).
* ``--to`` alone     -> in-place rewrite, but ``<file>.bak`` first, holding the
                       pre-edit bytes.
* ``-o`` into a missing directory -> non-zero exit, ``--to`` untouched, and the
                       directory is NOT auto-created (a typo'd path should not
                       quietly invent directories).
* ``extract -o``     -> untouched semantics (JSON path), guarded so a fix in the
                       ``apply`` branch cannot drift into it.
"""
from __future__ import annotations

import contextlib
import hashlib
import io
import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document                       # noqa: E402
from docx.shared import Pt, Twips                # noqa: E402

from docx_template import _extract, main as template_main  # noqa: E402
from _omml import needs_omml                  # noqa: E402  (single-source OMML gate)

# Deliberately different from the python-docx default (Letter, 1in margins) so
# "did the template land?" is decidable from the output alone.
TPL_PAGE_WIDTH = 11906      # A4
TPL_PAGE_HEIGHT = 16838
TPL_TOP_MARGIN = 1134       # 2.0 cm
TPL_LEFT_MARGIN = 2160      # 3.8 cm
TPL_FONT = "Courier New"
TPL_FONT_SIZE = 13


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def part_bytes(docx: Path, name: str) -> bytes:
    with zipfile.ZipFile(docx) as z:
        return z.read(name)


def run_cli(argv: list[str]) -> int:
    """Call the CLI in-process; ``sys.exit(1)`` paths become a return code."""
    saved = sys.argv
    try:
        sys.argv = ["docx_template.py", *argv]
        try:
            return template_main()
        except SystemExit as exc:             # argparse / sys.exit(1) branches
            return exc.code if isinstance(exc.code, int) else 1
    finally:
        sys.argv = saved


@pytest.fixture()
def fixtures(tmp_path: Path) -> dict:
    """A template source (distinctive layout) and a plain target document."""
    src = Document()
    src.add_paragraph("template source")
    sec = src.sections[0]
    sec.page_width = Twips(TPL_PAGE_WIDTH)
    sec.page_height = Twips(TPL_PAGE_HEIGHT)
    sec.top_margin = Twips(TPL_TOP_MARGIN)
    sec.left_margin = Twips(TPL_LEFT_MARGIN)
    normal = src.styles["Normal"]
    normal.font.name = TPL_FONT
    normal.font.size = Pt(TPL_FONT_SIZE)
    tpl_source = tmp_path / "tpl_source.docx"
    src.save(tpl_source)

    tgt = Document()
    tgt.add_paragraph("payload paragraph")
    tgt.add_paragraph("second line")
    target = tmp_path / "in.docx"
    tgt.save(target)

    tpl = _extract(tpl_source)
    assert tpl["page"]["margins"]["top"] == TPL_TOP_MARGIN       # _extract yields ints
    assert tpl["page"]["margins"]["top"] != _extract(target)["page"]["margins"]["top"]

    tpl_json = tmp_path / "tpl.json"
    tpl_json.write_text(json.dumps(tpl, ensure_ascii=False, indent=2), encoding="utf-8")
    return {"tmp": tmp_path, "tpl": tpl_json, "tpl_data": tpl,
            "in": target, "bak": Path(str(target) + ".bak")}


# ---------------------------------------------------------------------------
# 1. the load-bearing case: -o is written, --to keeps its exact bytes
# ---------------------------------------------------------------------------
def test_to_plus_output_writes_output_and_leaves_to_file_identical(fixtures):
    target, out = fixtures["in"], fixtures["tmp"] / "out.docx"
    before = sha256(target)

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target), "-o", str(out)])

    assert rc == 0
    assert out.is_file(), "-o target was not created (the P0-B bug)"
    assert sha256(target) == before, "--to source was rewritten; -o was ignored"
    assert target.stat().st_size == len(target.read_bytes())
    assert not fixtures["bak"].exists(), "no backup should be needed when -o is given"


# ---------------------------------------------------------------------------
# 2. the output really carries the template (not an empty file / a plain copy)
# ---------------------------------------------------------------------------
def test_output_contains_applied_template_content(fixtures):
    target, out = fixtures["in"], fixtures["tmp"] / "out.docx"
    before_bytes = target.read_bytes()

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target), "-o", str(out)])
    assert rc == 0

    # Style evidence: the template font exists in the copy, not in the original.
    assert TPL_FONT.encode() in part_bytes(out, "word/styles.xml")
    assert TPL_FONT.encode() not in part_bytes(target, "word/styles.xml")

    # Page-geometry evidence: the copy round-trips to the template's page dict.
    assert _extract(out)["page"] == fixtures["tpl_data"]["page"]

    # And the source is still literally the same file content.
    assert target.read_bytes() == before_bytes


# ---------------------------------------------------------------------------
# 3. no -o: in-place rewrite is allowed, but only behind a .bak
# ---------------------------------------------------------------------------
def test_in_place_apply_writes_backup_equal_to_pre_edit_bytes(fixtures, capsys):
    target, bak = fixtures["in"], fixtures["bak"]
    before = sha256(target)
    assert not bak.exists()

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target)])
    out = capsys.readouterr().out

    assert rc == 0
    assert bak.is_file(), "in-place apply must leave a .bak"
    assert sha256(bak) == before, ".bak must hold the pre-edit bytes"
    assert sha256(target) != before, "in-place apply must actually rewrite the file"
    assert _extract(target)["page"] == fixtures["tpl_data"]["page"]
    assert "backup" in out.lower(), f"stdout should announce the backup, got: {out!r}"


# ---------------------------------------------------------------------------
# 4. extract -o semantics are untouched (regression guard for the fix's blast radius)
# ---------------------------------------------------------------------------
def test_extract_output_flag_still_writes_json_path(fixtures):
    src, meta = fixtures["tmp"] / "tpl_source.docx", fixtures["tmp"] / "meta.json"
    before = sha256(src)

    rc = run_cli(["extract", str(src), "-o", str(meta)])

    assert rc == 0
    assert meta.is_file()
    assert sorted(json.loads(meta.read_text(encoding="utf-8"))) == ["defaults", "page", "styles"]
    assert sha256(src) == before, "extract must never touch the document it reads"
    assert not (src.parent / (src.stem + "_template.json")).exists(), \
        "-o was given, so no default-named json should appear"


# ---------------------------------------------------------------------------
# 5. boundary: -o pointing into a directory that does not exist
# ---------------------------------------------------------------------------
def test_output_into_missing_directory_fails_without_touching_source(fixtures, capsys):
    target = fixtures["in"]
    missing_dir = fixtures["tmp"] / "no_such_dir"
    out = missing_dir / "out.docx"
    before = sha256(target)

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target), "-o", str(out)])
    captured = capsys.readouterr()

    # Measured choice: fail loudly instead of auto-creating the directory, so a
    # typo'd path cannot silently scatter output somewhere unexpected.
    assert rc != 0, "a missing output directory must be an error, not a silent no-op"
    assert "does not exist" in captured.err
    assert not out.exists()
    assert not missing_dir.exists(), "the directory must not be created behind the user's back"
    assert sha256(target) == before, "a failed -o run must leave the --to file untouched"


# ---------------------------------------------------------------------------
# 6. edge: -o names the same file as --to -> treated as in-place, so .bak appears
# ---------------------------------------------------------------------------
def test_output_equal_to_to_path_falls_back_to_in_place_with_backup(fixtures):
    target, bak = fixtures["in"], fixtures["bak"]
    before = sha256(target)

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target), "-o", str(target)])

    assert rc == 0
    assert sha256(target) != before
    assert bak.is_file() and sha256(bak) == before


# ---------------------------------------------------------------------------
# 7. edge: an existing -o file is overwritten, the source still is not
# ---------------------------------------------------------------------------
def test_existing_output_file_is_overwritten(fixtures):
    target, out = fixtures["in"], fixtures["tmp"] / "out.docx"
    out.write_bytes(b"stale bytes, not a docx")
    before = sha256(target)

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target), "-o", str(out)])

    assert rc == 0
    assert out.stat().st_size > len(b"stale bytes, not a docx")
    assert TPL_FONT.encode() in part_bytes(out, "word/styles.xml")
    assert sha256(target) == before


# ---------------------------------------------------------------------------
# 8-12. --fix-formulas / --no-fix-formulas: two channels, two defaults, and
#       "both" must be a usage error instead of a silent merge (P6/FIX-C5D).
#
# The audit's complaint was that the two flags read like a +/- pair while
# actually writing two different dests:
#     --no-fix-formulas -> dest=fix_formulas  (marker channel, default ON)
#     --fix-formulas    -> dest=omml_formulas (OMML channel,   default OFF)
# Measured on the pre-fix code (2026-11-11): passing both exited 0, turned the
# marker channel off, and handed raw ``_xx_`` markers to docx_mathify -- which
# then failed on any document carrying markers, still with exit 0. What is
# asserted below is the post-fix contract: defaults unchanged, each single flag
# unchanged, both together exit 2 with nothing written.
# ---------------------------------------------------------------------------
MARKER_PARA = "water H_2_O marker"          # paired _x_ -> vertAlign subscript
LATEX_PARA = "结论：\\[ \\frac{a}{b} = c \\]"   # structural LaTeX -> OMML


@pytest.fixture()
def formula_docs(fixtures):
    """A target whose two paragraphs exercise the two channels separately.

    They are kept in *different* paragraphs on purpose: docx_mathify rewrites a
    paragraph it converts, so a marker and a formula in the same paragraph would
    make "did the marker channel run?" unobservable.
    """
    target = fixtures["tmp"] / "formulas.docx"
    doc = Document()
    doc.add_paragraph(MARKER_PARA)
    doc.add_paragraph(LATEX_PARA)
    doc.save(target)
    return {**fixtures, "target": target,
            "out": fixtures["tmp"] / "formulas_out.docx"}


def counts(docx: Path) -> tuple[int, int]:
    """(vertAlign occurrences, <m:oMath> occurrences) in document.xml."""
    xml = part_bytes(docx, "word/document.xml").decode("utf-8")
    return xml.count("vertAlign"), xml.count("<m:oMath")


def test_no_formula_flag_keeps_documented_defaults(formula_docs):
    """No flag: marker channel ON, OMML channel OFF -- exactly the old default."""
    out = formula_docs["out"]

    rc = run_cli(["apply", str(formula_docs["tpl"]),
                  "--to", str(formula_docs["target"]), "-o", str(out)])

    assert rc == 0
    vertalign, omath = counts(out)
    assert vertalign >= 1, "default marker channel stopped running"
    assert omath == 0, "the OMML channel is opt-in; it must not run on its own"


def test_no_fix_formulas_alone_only_disables_the_marker_channel(formula_docs):
    out = formula_docs["out"]

    rc = run_cli(["apply", str(formula_docs["tpl"]), "--to",
                  str(formula_docs["target"]), "-o", str(out), "--no-fix-formulas"])

    assert rc == 0
    vertalign, omath = counts(out)
    assert vertalign == 0, "--no-fix-formulas failed to disable the marker channel"
    assert omath == 0, "the marker flag must not switch the OMML channel on"


@needs_omml          # oMath 断言需要 Office；见 docs/reports/P6-FIX-C5D.md
def test_fix_formulas_alone_adds_omml_and_leaves_marker_channel_on(formula_docs):
    """The two channels coexist: opting into OMML does not disable the markers.

    ⚠️ 这条断言的是**两条通道同时可见**（`oMath>=1` 且 `vertAlign>=1`），所以
    必须挂 ``@needs_omml``：CI 四腿没有 Office、`MML2OMML.XSL` 不存在，
    ``docx_mathify`` 会降级/报错，`oMath` 恒为 0 —— 不挂标记就是必红。
    断言本身不降级：不能改成 `oMath>=0`，那等于什么都没测。
    """
    out = formula_docs["out"]

    rc = run_cli(["apply", str(formula_docs["tpl"]), "--to",
                  str(formula_docs["target"]), "-o", str(out), "--fix-formulas"])

    assert rc == 0
    vertalign, omath = counts(out)
    assert omath >= 1, "--fix-formulas did not convert the LaTeX paragraph to OMML"
    assert vertalign >= 1, \
        "--fix-formulas must not turn off the default marker channel"


@pytest.mark.parametrize("flags", [
    ["--fix-formulas", "--no-fix-formulas"],
    ["--no-fix-formulas", "--fix-formulas"],
])
def test_both_formula_flags_are_rejected_with_exit_2(formula_docs, capsys, flags):
    """Both at once: argparse usage error, exit 2, nothing written."""
    out = formula_docs["out"]

    rc = run_cli(["apply", str(formula_docs["tpl"]),
                  "--to", str(formula_docs["target"]), "-o", str(out), *flags])
    captured = capsys.readouterr()

    assert rc == 2, f"expected argparse's usage exit 2, got {rc}"
    assert "not allowed with argument" in captured.err, \
        f"the error must name the conflict, got: {captured.err!r}"
    assert not out.exists(), "a rejected invocation must not write a document"


def test_apply_help_states_the_two_channels_and_their_exclusivity():
    """--help is the one place a user can see both channels and the XOR."""
    saved = sys.argv
    buf = io.StringIO()
    try:
        sys.argv = ["docx_template.py", "apply", "--help"]
        try:
            with contextlib.redirect_stdout(buf):
                template_main()
        except SystemExit as exc:
            assert exc.code == 0
    finally:
        sys.argv = saved
    # argparse wraps help text at the terminal width, so compare on a flattened
    # copy: the claim being asserted is the wording, not where the line breaks.
    flat = " ".join(buf.getvalue().split())

    assert "[--no-fix-formulas | --fix-formulas]" in flat, \
        "usage line must show the flags as mutually exclusive"
    assert "Opt-in" in flat, "--fix-formulas must be labelled opt-in"
    assert "independent of the marker channel" in flat, \
        "--fix-formulas must be described as a separate channel"


# ---------------------------------------------------------------------------
# 13-16. --fix-formulas must NOT swallow docx_mathify's exit code (P8-D6 / A).
#
# The bug (measured 2026-11-11 on the pre-fix code): ``_convert_formulas`` called
# ``docx_mathify.main()`` and dropped the return value, so all three call sites
# lost it and ``main()`` ended with an unconditional ``return 0``. A document
# whose formula latex2mathml rejects therefore produced
#     stderr: error: 1 formula(s) failed to convert
#     exit:   0
# i.e. a half-converted file handed over as if it were finished. D-022 already
# defines mathify's code 3 as "written, but at least one formula did not
# convert"; that code has to survive the template layer.
#
# Gating, precisely: only the two cases that assert an ``m:oMath`` object are
# marked ``needs_omml``. The propagation cases above are *not* gated and must not
# be: they assert the exit code of a formula that fails to convert, which the
# template layer reports whether or not the XSL is present (measured: removing
# the marks leaves them passing on a machine with no Office). Gating them would
# throw away real coverage on every runner. The ``--no-fix-formulas`` control is
# likewise unmarked -- it never calls mathify.
# ---------------------------------------------------------------------------
BAD_FORMULA_PARA = r"坏公式 $x=\frac{a}{y^{%}}$ 结束"   # % as a superscript: latex2mathml rejects
GOOD_FORMULA_PARA = r"结论 $x=\frac{a}{b}$ 结束"        # legal: converts when XSL is present


@pytest.fixture()
def broken_formula_doc(fixtures):
    """A target carrying one formula the real chain cannot convert."""
    target = fixtures["tmp"] / "broken.docx"
    doc = Document()
    doc.add_paragraph(BAD_FORMULA_PARA)
    doc.save(target)
    return {**fixtures, "target": target,
            "out": fixtures["tmp"] / "broken_out.docx"}


def test_fix_formulas_propagates_exit_3_in_place(broken_formula_doc, capsys):
    """In-place ``--to`` + a bad formula: exit 3 AND the document is written.

    Both halves matter: 3 is the only "produced something + non-zero" exit
    (D-022), so a fix that returned 1 here would wrongly tell the user nothing
    landed, and one that returned 0 is the bug being fixed.
    """
    target = broken_formula_doc["target"]

    rc = run_cli(["apply", str(broken_formula_doc["tpl"]), "--to", str(target),
                  "--fix-formulas"])
    captured = capsys.readouterr()

    assert rc == 3, (
        f"模板吞掉了 mathify 的退出码 3，实际返回 {rc}；stdout={captured.out!r} "
        f"stderr={captured.err!r}")
    assert "failed to convert" in captured.err, \
        f"mathify 的失败诊断应当原样透出，实际 stderr={captured.err!r}"
    assert target.is_file(), "退 3 的语义是「文档已写出」，产物必须还在"
    # stdout 逐字契约：本单不许新增行
    assert "Applied template" in captured.out and "(in place)" in captured.out


def test_fix_formulas_propagates_exit_3_when_copying_to_output(broken_formula_doc, capsys):
    """The ``--to`` + ``-o`` call site must not swallow the code either."""
    target, out = broken_formula_doc["target"], broken_formula_doc["out"]
    before = sha256(target)

    rc = run_cli(["apply", str(broken_formula_doc["tpl"]), "--to", str(target),
                  "-o", str(out), "--fix-formulas"])
    captured = capsys.readouterr()

    assert rc == 3, (
        f"-o 分支吞掉了 mathify 的退出码 3，实际返回 {rc}；stderr={captured.err!r}")
    assert out.is_file(), "退 3 时输出文件必须存在（有产物 + 非零码）"
    assert sha256(target) == before, "-o 分支不许碰 --to 源文件"


def test_no_fix_formulas_keeps_exit_0_on_the_same_document(broken_formula_doc, capsys):
    """Control: without the OMML channel the very same document exits 0.

    This is what proves the 3 above comes from the mathify channel and not from
    the template layer suddenly refusing bad formulas on its own.
    """
    target = broken_formula_doc["target"]

    rc = run_cli(["apply", str(broken_formula_doc["tpl"]), "--to", str(target),
                  "--no-fix-formulas"])
    captured = capsys.readouterr()

    assert rc == 0, f"--no-fix-formulas 不跑公式链，退出码应保持 0，实际 {rc}"
    assert "failed to convert" not in captured.err, \
        "marker 通道不该报出公式转换失败"
    assert target.is_file()


@needs_omml
def test_fix_formulas_still_exits_0_when_every_formula_converts(fixtures, capsys):
    """Positive control: a clean document must still exit 0, not "always 3"."""
    target = fixtures["tmp"] / "clean.docx"
    doc = Document()
    doc.add_paragraph(GOOD_FORMULA_PARA)
    doc.save(target)

    rc = run_cli(["apply", str(fixtures["tpl"]), "--to", str(target),
                  "--fix-formulas"])
    captured = capsys.readouterr()

    assert rc == 0, (
        f"公式全部转换成功时退出码必须是 0，实际 {rc}；stderr={captured.err!r}")
    assert "failed to convert" not in captured.err
    vertalign, omath = counts(target)
    assert omath >= 1, "前置假设失效：合法公式没有被转成 OMML"


# ===========================================================================
# 输入错误契约（P9-C）
#
# 三个子命令各有各的「输入存在性」检查，且**都不是同一种写法**（实测
# 2026-11-11）：
#
#   extract  只查 `src.exists()`   → 缺文件 1 行 exit 1；目录 17 行 traceback；
#                                   非 zip 18 行 traceback（`_read_blobs` 炸）
#   fill     查模板 + values 两个路径，非 zip 由 PackageNotFoundError 兜成 1 行
#            exit 1（语义错：这属于「输入根本不是包」，不是「填不进去」）
#   apply    只查 `--to` 是否存在  → 非 zip **先** `shutil.copyfile` 写出
#            `<名>.bak` **再**崩，与 docx_mathify 的 `--inplace` 同型
#
# 本节的用例一律走**子进程**：本文件既有的 `run_cli()` 是 in-process
# （改写 `sys.argv` 后调 `template_main()`），拿得到返回码，却拿不到
# 「stderr 恰好一行」与「stdout 零字节」这两件事 —— 而那正是本单要钉的。
# ===========================================================================
def run_cli_proc(argv: list, cwd: Path | None = None) -> subprocess.CompletedProcess:
    """子进程跑 CLI。

    ``cwd`` 不能省：`extract` 的默认输出名是**相对路径**（``<stem>_template.json``），
    不指定工作目录就会落到跑 pytest 的那个目录（仓库根），把测试垃圾写进仓。
    """
    env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_template.py"), *map(str, argv)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
        cwd=str(cwd) if cwd else None)


def one_error_line(stderr: str) -> str:
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def plain_docx(tmp_path: Path, name: str = "plain.docx") -> Path:
    d = Document()
    d.add_paragraph("hello")
    p = tmp_path / name
    d.save(str(p))
    return p


def nonzip(tmp_path: Path, name: str = "renamed.docx") -> Path:
    p = tmp_path / name
    p.write_text("i am not a zip package", encoding="utf-8")
    return p


def test_extract_missing_path_is_a_usage_error(tmp_path: Path):
    """``extract`` 缺文件 → exit 2 + 单行（不再走 exit 1 的「处理不了」通道）。"""
    missing = tmp_path / "no-such.docx"
    proc = run_cli_proc(["extract", missing])
    assert proc.returncode == 2, proc.stderr
    line = one_error_line(proc.stderr)
    assert "no such file" in line
    assert str(missing) in line
    assert proc.stdout == ""


def test_extract_directory_is_a_usage_error(tmp_path: Path):
    """``extract`` 传目录 → exit 2 + 单行，而非 17 行 traceback。"""
    adir = tmp_path / "a-directory"
    adir.mkdir()
    proc = run_cli_proc(["extract", adir])
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a file: {adir}"
    assert proc.stdout == ""


def test_extract_non_zip_is_a_usage_error(tmp_path: Path):
    """``extract`` 非 zip → exit 2 + 单行，且不生成 ``<名>_template.json``。"""
    bogus = nonzip(tmp_path)
    proc = run_cli_proc(["extract", bogus])
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a ZIP package: {bogus}"
    assert proc.stdout == ""
    assert not (tmp_path / "renamed_template.json").exists()


def test_fill_non_zip_template_is_a_usage_error(tmp_path: Path):
    """``fill`` 的模板不是包 → exit 2（旧版是 exit 1 的 ``is not a .docx package``）。"""
    bogus = nonzip(tmp_path)
    values = tmp_path / "v.json"
    values.write_text("{}", encoding="utf-8")
    out = tmp_path / "out.docx"

    proc = run_cli_proc(["fill", bogus, values, out])
    assert proc.returncode == 2, proc.stderr
    assert one_error_line(proc.stderr) == f"error: not a ZIP package: {bogus}"
    assert proc.stdout == ""
    assert not out.exists()


def test_apply_non_zip_target_writes_no_backup(tmp_path: Path):
    """★ ``apply --to`` 指向非 zip → exit 2 且**一个文件都不生成**。

    旧版走到 ``dest is target`` 分支，先 ``shutil.copyfile(target, backup)``
    留下 ``renamed.docx.bak``，然后在 ``_apply_existing`` 上崩 —— 输入是垃圾
    却留下垃圾备份。
    """
    bogus = nonzip(tmp_path)
    tpl = tmp_path / "tpl.json"
    tpl.write_text(json.dumps(_extract(plain_docx(tmp_path))), encoding="utf-8")
    before = {p.name for p in tmp_path.iterdir()}

    proc = run_cli_proc(["apply", tpl, "--to", bogus])
    assert proc.returncode == 2, proc.stderr
    assert proc.stdout == ""
    after = {p.name for p in tmp_path.iterdir()}
    assert after == before, f"拒绝时写盘了：新增 {sorted(after - before)}"
    assert not (tmp_path / "renamed.docx.bak").exists()


def test_extract_valid_document_still_writes_json(tmp_path: Path):
    """合法文档 ``extract`` 仍是 exit 0 + 写出 JSON（加校验没把成功路径改坏）。"""
    src = plain_docx(tmp_path)
    proc = run_cli_proc(["extract", src], cwd=tmp_path)
    assert proc.returncode == 0, proc.stderr
    assert proc.stderr == ""
    out = tmp_path / "plain_template.json"
    assert out.is_file()
    payload = json.loads(out.read_text(encoding="utf-8"))
    assert set(payload) == {"page", "defaults", "styles"}