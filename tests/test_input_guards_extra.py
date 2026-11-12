# -*- coding: utf-8 -*-
"""``docx_comments.py`` / ``docx_revisions.py`` 的**输入守卫**（P9-D）。

## 为什么单独开一个文件

这两个脚本在全仓**没有专项测试文件**（其余用例零散挂在 ``test_docx_skill.py``
的功能族里，那是 e2e 幸福路径，不测退出码契约）。本文件只装它们俩的输入守卫，
与 P9-C 的 ``test_validate_errors.py`` / ``test_numbering_errors.py`` 同形。
``merge_docs`` 与 ``insert_formula`` 各有专项文件（``test_merge.py`` /
``test_insert_formula.py``），用例落在各自文件里，不再重复建文件。

## 被锁的契约

    exit 2  你给的东西根本不是一个可读的 .docx（stdout **空**，stderr **单行** ``error: …``）
    exit 0  文档正常（输出 JSON）
    exit 1  文档是真的、但请求的东西不在里面（``--target`` / ``--id`` 未命中，输出 JSON）

四档坏输入各锁一次：缺文件 / 是目录 / 非 zip / zip 但非 Word 包。

## ⚠️ 两条方法论（本单的重点，别退回旧形态）

**1. argv 必须是「能真正走到 ``_open_document`` 的最短合法调用」。**

旧版真正崩在哪，只能用完整 argv 才看得见（实测 2026-11-11）：

```
docx_comments.py NOPE.docx            → rc=2   ← 假绿：argparse 报 invalid choice
docx_comments.py list NOPE.docx       → rc=1 + 19 行裸 traceback   ← 真相
```

所以本文件**每一条**坏输入用例都带子命令，且额外断言 stderr 点出了那个坏路径
——argparse 阶段的退出不会提路径，这个断言就是「没走 argparse 假绿」的判别器。

**2. 校验必须覆盖**每一个**会打开文档的子命令**，不是只修 ``list``。**

两个脚本的校验都写在 ``main()`` 里「拿 args → 打开文档」的**唯一出口**上，
所以 ``add`` / ``delete`` / ``accept-all`` / ``reject-all`` / ``accept`` / ``reject``
与 ``list`` 共用同一条路。用例对**每个子命令**都跑一遍四档坏输入：只测 ``list``
的话，将来有人把某条子命令改成自己 `Document()` 就发现不了。

## 为什么**不**挂 ``@needs_omml``

本文件只用 ``python-docx`` 造一个正常文档 + ``zipfile``/纯文本造坏输入，不碰
``math_omml`` / ``MML2OMML.XSL`` / Office COM。无 Office 的 CI runner 上必须
**pass**（不是 skip）。实测：``MML2OMML_XSL`` 指到不存在路径后本文件全 pass。

Run:  python -m pytest tests/test_input_guards_extra.py -q
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
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))

docx = pytest.importorskip("docx", reason="需要 python-docx 造 fixture")

COMMENTS = SCRIPTS / "docx_comments.py"
REVISIONS = SCRIPTS / "docx_revisions.py"


# ---------------------------------------------------------------- helpers
def run(script: Path, *argv) -> subprocess.CompletedProcess:
    """以真子进程跑 CLI（退出码是进程级契约，不能 in-process 测）。"""
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(script), *map(str, argv)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )


def one_error_line(stderr: str) -> str:
    """断言 stderr **恰好一行**（非空行计），返回该行。"""
    lines = [ln for ln in stderr.strip().splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr}"
    return lines[0]


def healthy_docx(tmp_path: Path) -> Path:
    d = docx.Document()
    d.add_paragraph("alpha one")
    d.add_paragraph("Q3 revenue grew")
    p = tmp_path / "healthy.docx"
    d.save(str(p))
    return p


# ---------------------------------------------------------------- 四档坏输入
@pytest.fixture(params=["missing", "directory", "non-zip", "zip-not-word"])
def bad_input(request, tmp_path: Path) -> tuple[Path, str]:
    """造一档坏输入，返回 ``(路径, 该档应当在 stderr 里出现的措辞)``。"""
    kind = request.param
    if kind == "missing":
        p = tmp_path / "no-such.docx"      # 刻意不创建
        return p, "no such file"
    if kind == "directory":
        p = tmp_path / "a-directory.docx"  # 目录顶着 .docx 的名字（shell glob 常见形态）
        p.mkdir()
        return p, "not a file"
    if kind == "non-zip":
        p = tmp_path / "renamed.docx"
        p.write_text("this is not a zip file", encoding="utf-8")
        return p, "not a ZIP package"
    # 合法 zip，但没有 [Content_Types].xml → is_zipfile 放行，Document() 才炸
    p = tmp_path / "bare-zip.docx"
    with zipfile.ZipFile(p, "w") as zf:
        zf.writestr("foo.txt", "bar")
    return p, "not a Word package"


# ===========================================================================
# docx_comments —— 三个子命令 × 四档坏输入
# ===========================================================================
COMMENTS_SUBCOMMANDS = [
    ("list", lambda p: ["list", p]),
    ("add", lambda p: ["add", p, "--target", "alpha", "--text", "note"]),
    ("delete", lambda p: ["delete", p, "--id", "0"]),
]


@pytest.mark.parametrize("cmd,build", COMMENTS_SUBCOMMANDS,
                         ids=[c for c, _ in COMMENTS_SUBCOMMANDS])
def test_comments_bad_input_is_exit_2_for_every_subcommand(cmd, build, bad_input):
    """``docx_comments`` 的**每个**子命令：四档坏输入 → exit 2 + 单行 + stdout 空。

    只测 ``list`` 是不够的：校验虽然在共用出口上，但用例是唯一能防「某条子命令
    被改成自己 `Document()`」的东西。三条子命令 × 四档坏输入 = 12 次真进程调用，
    全部必须在**同一条出口**上退 2。

    ``stderr`` 里必须出现该档的**具体措辞**（``no such file`` / ``not a file`` /
    ``not a ZIP package`` / ``not a Word package``），且必须带那个坏路径：

    * 措辞断言锁住「四档没被压成同一句含糊报错」；
    * 路径断言是**假绿判别器** —— argparse 阶段的 ``invalid choice`` 不会提路径。
    """
    path, wording = bad_input
    proc = run(COMMENTS, *build(str(path)))

    assert proc.returncode == 2, (
        f"`{cmd}` 收到坏输入必须退 2：rc={proc.returncode}\n"
        f"stdout={proc.stdout!r}\nstderr={proc.stderr!r}")
    line = one_error_line(proc.stderr)
    assert line.startswith("error: "), f"stderr 不是 `error: ` 前缀：{line!r}"
    assert wording in line, f"报错没指出坏在哪（期望 {wording!r}）：{line!r}"
    assert str(path) in line, (
        f"报错没提那个坏路径 —— 这条是 argparse 假绿的判别器：{line!r}")
    assert proc.stdout == "", (
        f"失败路径不该有 stdout（成功与 `--id` 未命中都会打 JSON，"
        f"空 stdout 不能被读成空 JSON）：{proc.stdout!r}")


def test_comments_valid_document_still_works(tmp_path: Path):
    """合法对照：健康文档 → exit 0 + stdout 是带 ``ok:true`` 的 JSON。

    防「一律报错」的假修：校验加过头时这条先红。
    """
    src = healthy_docx(tmp_path)
    proc = run(COMMENTS, "list", src)

    assert proc.returncode == 0, f"rc={proc.returncode} stderr={proc.stderr!r}"
    assert proc.stderr == "", f"成功路径不该有 stderr：{proc.stderr!r}"
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert rep["comments"] == []


def test_comments_anchor_not_found_is_still_exit_1(tmp_path: Path):
    """**不许顺手改掉的那条**：文档真实、但 ``--target`` 找不到 → 仍是 exit 1 + JSON。

    这是「文档是真的、但你要的东西不在里面」，与「输入的路径根本不对」是两件事。
    把这条也提到 2 就等于抹掉这个区别 —— 本单只该动输入档，不动这一档。
    """
    src = healthy_docx(tmp_path)
    proc = run(COMMENTS, "add", src, "--target", "no such anchor",
               "--text", "note")

    assert proc.returncode == 1, (
        f"锚点未命中仍是 exit 1（运行期语义，不是用法错误）："
        f"rc={proc.returncode} stderr={proc.stderr!r}")
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    assert "target not found" in rep["error"]


# ===========================================================================
# docx_revisions —— 五个子命令 × 四档坏输入
# ===========================================================================
REVISIONS_SUBCOMMANDS = [
    ("list", lambda p: ["list", p]),
    ("accept-all", lambda p: ["accept-all", p]),
    ("reject-all", lambda p: ["reject-all", p]),
    ("accept", lambda p: ["accept", p, "--id", "1"]),
    ("reject", lambda p: ["reject", p, "--id", "1"]),
]


@pytest.mark.parametrize("cmd,build", REVISIONS_SUBCOMMANDS,
                         ids=[c for c, _ in REVISIONS_SUBCOMMANDS])
def test_revisions_bad_input_is_exit_2_for_every_subcommand(cmd, build, bad_input):
    """``docx_revisions`` 的**每个**子命令：四档坏输入 → exit 2 + 单行 + stdout 空。

    五个子命令里有三个是**写盘**子命令（``accept-all`` / ``reject-all`` /
    ``accept`` / ``reject`` 都会 ``doc.save()``）。它们在坏输入上退 2 还不算完，
    关键是**零写盘** —— 那一条由 ``test_merge.py`` 同族的零写盘用例覆盖（合并
    场景更危险：它会创建输出文件）。这里只锁出口契约：共 5 × 4 = 20 次真进程调用。
    """
    path, wording = bad_input
    proc = run(REVISIONS, *build(str(path)))

    assert proc.returncode == 2, (
        f"`{cmd}` 收到坏输入必须退 2：rc={proc.returncode}\n"
        f"stdout={proc.stdout!r}\nstderr={proc.stderr!r}")
    line = one_error_line(proc.stderr)
    assert line.startswith("error: "), f"stderr 不是 `error: ` 前缀：{line!r}"
    assert wording in line, f"报错没指出坏在哪（期望 {wording!r}）：{line!r}"
    assert str(path) in line, (
        f"报错没提那个坏路径 —— 这条是 argparse 假绿的判别器：{line!r}")
    assert proc.stdout == "", f"失败路径不该有 stdout：{proc.stdout!r}"


def test_revisions_bad_input_writes_nothing(tmp_path: Path):
    """坏输入 + ``-o``：**输出文件不许被创建**（写盘子命令的零写盘实证）。

    ``accept-all bad.docx -o out.docx`` 若在校验之前就 ``doc.save(out)``，会留下
    一个半成品输出文件 —— 那比崩掉更糟：调用方的下游会把它当成合并结果读。
    """
    out = tmp_path / "out.docx"
    missing = tmp_path / "no-such.docx"

    proc = run(REVISIONS, "accept-all", missing, "-o", out)

    assert proc.returncode == 2, f"rc={proc.returncode} stderr={proc.stderr!r}"
    assert not out.exists(), "坏输入路径仍然创建了输出文件（半成品落盘）"


def test_revisions_valid_document_still_works(tmp_path: Path):
    """合法对照：健康文档（零修订）→ exit 0 + ``ok:true`` + 空 revisions。

    注意这条与「坏输入」的区分点：**文档真实但没有修订是成功**（rc=0），
    不是错误。校验不许把这种情况也拒掉。
    """
    src = healthy_docx(tmp_path)
    proc = run(REVISIONS, "list", src)

    assert proc.returncode == 0, f"rc={proc.returncode} stderr={proc.stderr!r}"
    assert proc.stderr == "", f"成功路径不该有 stderr：{proc.stderr!r}"
    rep = json.loads(proc.stdout)
    assert rep["ok"] is True
    assert rep["revisions"] == []


def test_revisions_unknown_id_is_still_exit_1(tmp_path: Path):
    """**不许顺手改掉的那条**：文档真实、但 ``--id`` 不存在 → 仍是 exit 1 + JSON。"""
    src = healthy_docx(tmp_path)
    proc = run(REVISIONS, "accept", src, "--id", "999")

    assert proc.returncode == 1, (
        f"不存在的 id 仍是 exit 1（运行期语义，不是用法错误）："
        f"rc={proc.returncode} stderr={proc.stderr!r}")
    rep = json.loads(proc.stdout)
    assert rep["ok"] is False
    assert "no revision with id 999" in rep["error"]