# -*- coding: utf-8 -*-
r"""`insert_formula` 退出码分派 + stderr 契约（P28-A）。

## 这一组修的是什么

`insert_formula.py` 原来把三类性质完全不同的失败压成一条：

    except (RuntimeError, FileNotFoundError):
        print(f"error: {exc}", file=sys.stderr)
        return 2

于是「公式内容转不了」和「你参数写错了」共用一个码，而 D-024 对 `2` 的判据是
「拒绝先于任何读写：不写文件、stdout 为空、**stderr 恰好一行**」—— 守卫的四段多行
文案直接把后半条也破掉。P28-A 改成按 `math_omml` 的**异常类型**白名单分派：

| 失败性质 | 异常类型 | 退出码 | stderr 契约 |
|---|---|---|---|
| 源文本含 OMML 表达不了的控制序列（层 1 硬拒 / 层 2 残留） | `LatexGuardError` | **1** | 多行合法（四段文案） |
| `latex2mathml` 语法错 / XSL 应用失败 | `LatexSyntaxError` | **1** | 多行合法 |
| OMML 结构非法 / XSLT 没产出 | `OMMLStructureError` | **1** | 多行合法 |
| 缺 `latex2mathml` / XSL 找不到或加载不了 | `MathEnvError` / `FileNotFoundError` | **2** | **恰好一行** |
| 用法错（互斥 flag、空 marker、坏目录） | 不看异常，直接退 | **2** | **恰好一行** |


## 为什么「rc 对」本身不足以证明修好了

未捕获的异常在 CPython 里也是 **exit 1**。所以内容类用例只断 `rc == 1` 的话，一个
「异常直接掀栈」的实现照样能过 —— 那正是 P9 那轮抓到的形态。因此每条内容类用例都
同时断：`error:` 前缀、**无 `Traceback`**、stdout 为空、`document.xml` SHA 不变。

## 标记口径（P27 四腿全红的原路，别重踩）

* **层 1 拒收**（`\eqno` / `\textwidth`）发生在 `assert_no_unsupported_commands`，
  位于 `lconv.convert` 与 `xslt()` 之前 → 只要 `latex2mathml`，**标 `needs_l2m`**。
* **层 2 残留**（`\zzz`）要真转出 OMML 才扫得到 → **必须 `needs_omml`**。
* **环境类坏 XSL**：`etree.parse` 在 `xslt()` 里，`latex2mathml` 之前不用 XSL、之后
  才用到我们**指定**的坏文件 → **标 `needs_l2m`**（不依赖 Office 的 MML2OMML.XSL）。
* 用法错与 AST 防线：裸跑，不挂门控。

## 两条实测过的坑

1. **`--dump` 不需要 unpacked_dir 存在**（语义就是只打印），但内容类用例要断言
   「文档零改动」，所以一律给真解包目录 + `--after` 锚点走**写入分支**：
   失败发生在写盘之前才有意义，只跑 `--dump` 证明不了没写。
2. **坏 XSL 现场造**（`tmp_path` 写一段非 XML 文本），不依赖 `work/` 里的固定文件；
   并显式从子进程 env 里清掉继承来的 `MML2OMML_XSL`，否则本机 Office 会把它盖掉。

Run:  python -m pytest tests/test_insert_formula_exit_codes.py -q
"""
from __future__ import annotations

import ast
import hashlib
import os
import subprocess
import sys
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parent.parent
SCRIPTS = SKILL / "scripts"
sys.path.insert(0, str(SCRIPTS))

from _omml import needs_omml  # noqa: E402

# python-docx 是硬依赖（造 fixture 用），缺了直接失败而非 skip
docx = pytest.importorskip("docx", reason="需要 python-docx 造 fixture")
from docx_pack import unpack  # noqa: E402


def _has_latex2mathml() -> bool:
    """层 1 只依赖 latex2mathml（在它之前就抛），不需要 Office。"""
    try:
        import latex2mathml  # noqa: F401
    except Exception:
        return False
    return True


needs_l2m = pytest.mark.skipif(
    not _has_latex2mathml(), reason="needs latex2mathml")


# ---------------------------------------------------------------- helpers
def make_unpacked(tmp_path: Path, paragraphs: list) -> Path:
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


def run_insert(unpacked: Path, *args, env_extra: dict | None = None):
    """以子进程跑 ``insert_formula.py``（CLI 契约要按真进程验）。

    显式清掉继承来的 ``MML2OMML_XSL``：本机装了 Office 时它可能是空的、也可能是
    上一个用例留下的值，任何一种都会让「坏 XSL」用例测到别的东西。
    """
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8")
    env.pop("MML2OMML_XSL", None)
    if env_extra:
        env.update(env_extra)
    return subprocess.run(
        [sys.executable, str(SCRIPTS / "insert_formula.py"), str(unpacked),
         *map(str, args)],
        capture_output=True, env=env, check=False,
    )


def stderr_text(proc) -> str:
    return proc.stderr.decode("utf-8", "replace")


def stderr_lines(proc) -> list:
    return [ln for ln in stderr_text(proc).strip().splitlines() if ln.strip()]


def assert_single_line_error(proc, ctx: str) -> str:
    """2 档的统一判据：rc=2、stderr 恰好一行、``error:`` 前缀、无 traceback。"""
    err = stderr_text(proc)
    assert proc.returncode == 2, (
        "%s 必须退 2（用法 / 环境错）：rc=%s stdout=%r stderr=%r"
        % (ctx, proc.returncode, proc.stdout[:200], err[:400]))
    assert "Traceback (most recent call last)" not in err, (
        "%s 仍是裸 traceback（契约要求单行 error）：%r" % (ctx, err[:600]))
    lines = stderr_lines(proc)
    assert len(lines) == 1, (
        "%s 的 stderr 必须恰好一行，实际 %d 行：%r" % (ctx, len(lines), err[:600]))
    assert lines[0].startswith("error: "), (
        "%s 的 stderr 不是 error 前缀：%r" % (ctx, lines[0]))
    assert proc.stdout == b"", (
        "%s 的失败路径不许有 stdout：%r" % (ctx, proc.stdout[:200]))
    return lines[0]


# ===========================================================================
# 1) 内容类 → rc=1（守卫拒收 / 语法错 / 结构非法），且文档零字节改动
# ===========================================================================
# 三条覆盖两个不同的层，不能只用一条凑数：
#   * \eqno / \textwidth 被**层 1** 在 lconv.convert 之前拦下 → 只需 latex2mathml；
#   * \zzz 不在层 1 的表里，只能由**层 2**（扫成品 m:t）抓到 → 必须真转 OMML。
# 标反会把 CI 上本可真跑的用例全 skip 掉（P27 首轮四腿全红的原路）。
CONTENT_CASES = [
    pytest.param(r"E=mc^2 \eqno{3-1}", r"\eqno",
                 marks=needs_l2m, id="eqno-layer1"),
    pytest.param(r"w=0.5\textwidth", r"\textwidth",
                 marks=needs_l2m, id="textwidth-layer1"),
    pytest.param(r"a \zzz b", r"\zzz",
                 marks=needs_omml, id="zzz-layer2"),
]


@pytest.mark.parametrize("formula, command", CONTENT_CASES)
def test_content_failure_is_exit_1_and_document_untouched(
        formula, command, tmp_path: Path):
    r"""内容转不了 → **rc=1**（不是 2）+ stderr 点名命令 + 文档一个字节不动。

    走**写入分支**（给真解包目录 + ``--after`` 锚点）：只有这条路径才能证明
    「拒绝先于任何写盘」—— ``--dump`` 本来就不写文件，用它证明不了。

    五层断言缺一不可：

    * ``rc == 1``（``!= 0`` 不够；未捕获异常也是 1，所以还要看下面几条）；
    * stderr 以 ``error: `` 开头 —— 未捕获异常的第一行是 ``Traceback``，被这条挡下；
    * stderr 里**没有** ``Traceback`` —— 显式钉「异常被分派，不是掀栈」；
    * stderr 点名具体命令 —— 只有守卫的文案才有，证明拦截来自守卫而非上游碰巧报错；
    * ``document.xml`` SHA 不变；stdout 为空（没走到打印分支）。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])
    sha_before = doc_sha(up)

    proc = run_insert(up, formula, "--after", "marker here")
    err = stderr_text(proc)

    assert proc.returncode == 1, (
        "内容类失败必须退 1（输入合法但处理不了）：rc=%s stdout=%r stderr=%r"
        % (proc.returncode, proc.stdout[:200], err[:400]))
    assert "Traceback (most recent call last)" not in err, (
        "仍然是裸 traceback —— 异常没被分派到 1 档：%r" % (err[:600],))
    assert err.startswith("error: "), (
        "内容类失败的第一行不是 error 前缀：%r" % (err[:300],))
    assert command in err, (
        "报错没点名 %r —— 说明拦截不是来自守卫：%r" % (command, err[:300]))
    assert proc.stdout == b"", (
        "内容类失败不许有 stdout：%r" % (proc.stdout[:200],))

    assert doc_sha(up) == sha_before, (
        "失败路径改写了 document.xml：改前 %s → 改后 %s"
        % (sha_before[:12], doc_sha(up)[:12]))


# ===========================================================================
# 2) 环境类 → rc=2 且 stderr 恰好一行
# ===========================================================================
@needs_l2m
def test_probe5_bad_xsl_is_exit_2_single_line(tmp_path: Path):
    r"""**P28 探针 5 回归**：XSL 存在但不是合法 XML → rc=2 + 单行，不许掀栈。

    ## 缺陷形态（改之前实测，2026-11-26）

    ``xslt()`` 里的 ``etree.XSLT(etree.parse(...))`` 完全没有包装，坏 XSL 让
    ``lxml.etree.XMLSyntaxError`` 直接掀 **29 行** traceback、rc=1 —— 既不是单行，
    也不是 2（用户完全看不出该去修环境）。

    ## 断言为什么这么细

    这条同时是「``xslt()`` 的包装被删掉」的回归：去掉包装后 rc 会变回 1、行数变
    >=20，两条断言各自都会红。另外断「消息里带坏 XSL 的路径」—— 光看 rc=2 + 单行
    挡不住「换个原因退 2」（例如输入校验抢先在前面报错）。
    """
    bad = tmp_path / "bad.xsl"
    bad.write_text("not xml at all <<<", encoding="utf-8")
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])
    sha_before = doc_sha(up)

    proc = run_insert(up, "E=mc^2", "--after", "marker here",
                      env_extra={"MML2OMML_XSL": str(bad)})
    line = assert_single_line_error(proc, "[bad-xsl]")

    assert str(bad) in line, (
        "报错没提那个坏 XSL 文件 —— 换个原因退 2 也能混过这条：%r" % (line,))
    assert "无法加载" in line, (
        "报错没说清是 XSL 加载失败：%r" % (line,))
    assert doc_sha(up) == sha_before, "坏 XSL 路径改写了 document.xml"


@needs_l2m
def test_missing_xsl_is_exit_2_single_line(tmp_path: Path):
    r"""``MML2OMML_XSL`` 指向不存在的路径 → rc=2 + 单行（探针 4 的合规项，别回归）。

    这条走的是 ``find_xslt()`` 的 ``FileNotFoundError``，与坏 XSL 不同源：它必须在
    ``xslt()`` 的包装**之外**抛出（包进去就变成 ``MathEnvError``，文案与语义都变了）。
    改完仍要求 rc=2 + 单行 —— 这就是「包装不许吞掉 FileNotFoundError」的判据。
    """
    missing = tmp_path / "no-such.xsl"
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])
    sha_before = doc_sha(up)

    proc = run_insert(up, "E=mc^2", "--after", "marker here",
                      env_extra={"MML2OMML_XSL": str(missing)})
    line = assert_single_line_error(proc, "[missing-xsl]")

    assert str(missing) in line, (
        "报错没提那个不存在的 XSL 路径：%r" % (line,))
    assert doc_sha(up) == sha_before, "缺 XSL 路径改写了 document.xml"


USAGE_CASES = [
    pytest.param(("E=mc^2", "--after", "marker here", "--dump"),
                 id="dump-plus-after"),
    pytest.param(("E=mc^2", "--after", ""), id="empty-marker"),
    pytest.param(("E=mc^2", "--after", "x", "--before", "y"),
                 id="after-plus-before"),
]


@pytest.mark.parametrize("argv", USAGE_CASES)
def test_usage_errors_are_exit_2_single_line(argv, tmp_path: Path):
    r"""用法错 → rc=2 + 单行（D-024 原义，P28-A 不动这一档）。

    这一组是**对照**：分派改成白名单之后，既有 3 处 ``rc == 2`` 断言所覆盖的形态
    必须原样保持 —— 只有内容类降到 1，用法类一个都不许动。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])
    sha_before = doc_sha(up)

    proc = run_insert(up, *argv)
    assert_single_line_error(proc, "usage %r" % (argv,))
    assert doc_sha(up) == sha_before, "用法错误改写了 document.xml"


# ===========================================================================
# 3) 守卫文案完整（1 档没有单行约束 —— 四段文案必须原样留着）
# ===========================================================================
@needs_l2m
def test_guard_message_keeps_all_four_sections(tmp_path: Path):
    r"""内容类拒收时 stderr 里四段文案（症状 / 成因 / 替代写法 / 怎么做）都要在。

    判据是「1 档没有单行约束」这条裁定的**落地证据**：如果哪天有人为了迁就
    「stderr 恰好一行」去砍守卫文案，这条会红。同时断行数 >= 4 —— 只要有人把
    四段压成一句，即使关键词还在也会被行数判据抓到。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])
    proc = run_insert(up, r"E=mc^2 \eqno{3-1}", "--after", "marker here")
    err = stderr_text(proc)

    assert proc.returncode == 1, "守卫拒收退 1：rc=%s stderr=%r" % (
        proc.returncode, err[:300])
    for frag in ("症状：", "成因：", "替代写法：", "怎么做："):
        assert frag in err, "守卫文案缺少「%s」段：%r" % (frag, err[:600])
    assert len(stderr_lines(proc)) >= 4, (
        "四段文案被压成 %d 行：%r" % (len(stderr_lines(proc)), err[:600]))


# ===========================================================================
# 4) AST 防线：math_omml.py 里不许再有裸 raise RuntimeError
# ===========================================================================
def test_math_omml_has_no_bare_runtimeerror_raise():
    r"""结构性防线：``raise RuntimeError(...)`` 命中数必须是 **0**。

    ## 为什么这条必须存在（而不是靠人记住）

    病根是「新加一个抛点忘了分类，静默落回 2 冒充用法错」。类型分层本身治不了这个
    复发模式 —— 只有「裸 RuntimeError 一律红」这条机器判据能治。P28-A 落地时原有
    **9 处**全部换成子类，另在 ``xslt()`` 新增 1 处 ``MathEnvError``。

    报错信息里带行号，方便直接跳过去改（``ast`` 的 lineno 是 1-based）。
    """
    src = (SCRIPTS / "math_omml.py").read_text(encoding="utf-8")
    tree = ast.parse(src)

    offenders = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Raise) or node.exc is None:
            continue
        exc = node.exc
        if isinstance(exc, ast.Call):
            func = exc.func
            name = getattr(func, "id", None) or getattr(func, "attr", None)
            if name == "RuntimeError":
                offenders.append(node.lineno)
        elif isinstance(exc, ast.Name) and exc.id == "RuntimeError":
            offenders.append(node.lineno)

    assert offenders == [], (
        "math_omml.py 仍有裸 raise RuntimeError(...)（行号 %s）—— 每个抛点都必须"
        "落进 MathOmmlError 的子类，否则 insert_formula 会把它错判成退出码 2"
        "（用法错）。" % (offenders,))


def test_exception_hierarchy_keeps_runtimeerror_compatibility():
    r"""五个新类都必须继承 ``RuntimeError`` —— 向后兼容的唯一保证。

    ``test_math_latex_guard.py`` / ``test_math_omml_safe.py`` 里既有
    ``pytest.raises(RuntimeError)``，``docx_mathify`` / ``docx_create`` 里既有
    ``except Exception``：只要这五类的根仍是 ``RuntimeError``，它们的行为一字不变。
    这条把「分层不许破坏兼容」钉成机器判据，而不是写在报告里的一句话。
    """
    import math_omml

    names = ("MathOmmlError", "LatexGuardError", "LatexSyntaxError",
             "OMMLStructureError", "MathEnvError")
    for name in names:
        cls = getattr(math_omml, name)
        assert issubclass(cls, RuntimeError), "%s 不继承 RuntimeError" % name
        if name != "MathOmmlError":
            assert issubclass(cls, math_omml.MathOmmlError), (
                "%s 不是 MathOmmlError 的子类（层级被测塌了）" % name)


# ===========================================================================
# 5) 合法公式不误伤（防「一律报错」假修）
# ===========================================================================
@needs_omml
def test_valid_formula_still_exits_0(tmp_path: Path):
    r"""合法公式 → rc=0 + stdout 是可解析的 ``<w:p>``（``--dump`` 契约）。

    分派改成白名单后最容易出的假修是「宽了」：把内容类判成环境类、或者反过来一律
    报错。这条是唯一的正向对照，必须在。
    """
    up = make_unpacked(tmp_path, ["alpha one", "marker here", "omega three"])

    proc = run_insert(up, "E=mc^2", "--dump")
    assert proc.returncode == 0, (
        "合法公式必须退 0：rc=%s stderr=%r"
        % (proc.returncode, stderr_text(proc)[:300]))
    out = proc.stdout.decode("utf-8", "replace")
    assert out.startswith("<w:p"), "stdout 不是 w:p：%r" % (out[:200],)
    assert "<m:oMath" in out, "stdout 里没有 m:oMath：%r" % (out[:200],)
