# -*- coding: utf-8 -*-
r"""`docx_edit char-format` —— 字符级格式一等能力，以及 `docx_read --runs` 的回读。

## 这个文件存在的理由（真实事故，非臆想）

用户要在既有 Word 文档里把某些字标成 **深品红高亮**。手工 `docx_edit.py` 没有
任何 run 级格式子命令，于是他绕过工具直接改 XML：

```python
rPr = r._element.get_or_add_rPr()
hl = rPr.find(qn("w:highlight"))     # ← rPr 里没有 highlight 时返回 None
hl.set(qn("w:val"), "darkMagenta")   # ← AttributeError
```

这段代码有三个各自独立的坑，本文件逐个钉住：

1. **`find` 返回 `None`** —— `w:rPr` 是「有则用、无则建」，`get_or_add_rPr()`
   只保证 rPr 存在，不保证它里面有 `w:highlight`。空 rPr 上 `find` 给 `None`，
   `.set()` 直接 `AttributeError`（`work/p19c/probe1.py` 实测 F 段）。
2. **`append` 到 rPr 末尾** —— `w:rPr` 是 `xsd:sequence`，`w:highlight` 必须位于
   `w:szCs` 之后、`w:u` 之前。append 得到的是**合法 XML**、
   `docx_validate` 返 `ok:true`、**Word 静默忽略**（方法论 31）。
   probe1 实测：`['rFonts', 'b', 'sz', 'u', 'highlight']` —— highlight 排在 u 后面，错的。
3. **取值拼错不报错** —— `w:val` 是 `ST_HighlightColor` 枚举；写 `darkMagento`
   不抛异常、不显示高亮、没人告诉用户。

第 4 个坑是 python-docx 自身的：`WD_COLOR_INDEX` 只有 18 个成员，
**没有 `DARK_MAGENTA`**，`r.font.highlight_color = "darkMagenta"` 直接
`ValueError`（probe1 实测 E 段）。所以实现走 `qn()` + `OxmlElement` 自建校验，
`font.highlight_color` 不作唯一实现。

## 本文件的断言方式

「设了高亮」不靠 `in rPr` 断言 —— 那正是坑 2 能骗过 `docx_validate` 的原因。
这里断言的是 **兄弟元素序列**：`w:highlight` 的下标必须大于 `w:sz`/`w:szCs`
的下标、小于 `w:u` 的下标。位置错一位，测试就红。

纯 XML 层，**不挂 `needs_word` / `needs_omml` 门禁**：不读 Office、不读
`MML2OMML.XSL`，CI runner 与本机同结论。

Run:  python -m pytest tests/test_char_format.py -q
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.oxml.ns import qn
from docx.shared import Pt

SCRIPTS = Path(__file__).resolve().parent.parent / "scripts"
sys.path.insert(0, str(SCRIPTS))

EDIT = SCRIPTS / "docx_edit.py"
READ = SCRIPTS / "docx_read.py"


def run_cli(script: Path, *args):
    """Run a CLI and return ``(rc, stdout, stderr)``."""
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(script), *[str(a) for a in args]],
        capture_output=True, text=True, encoding="utf-8", errors="replace")
    return proc.returncode, proc.stdout, proc.stderr


def one_error_line(stderr: str) -> str:
    """The refusal contract: exactly one non-empty line, prefixed ``error:``."""
    lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(lines) == 1, f"expected one line on stderr, got {len(lines)}: {stderr!r}"
    assert lines[0].startswith("error:"), lines[0]
    return lines[0]


def rpr_tags(r_el) -> list:
    """Child local names of one run's ``w:rPr``, in document order."""
    rpr = r_el.find(qn("w:rPr"))
    assert rpr is not None, "run has no w:rPr after the command"
    return [child.tag.split("}")[1] for child in rpr]


def highlight_val(r_el):
    """``w:highlight/@w:val`` of one run, or ``None`` when it carries none.

    Written this way on purpose: ``r_el.find(qn("w:rPr")).find(...)`` dies with
    ``AttributeError`` on a run that never had any ``w:rPr`` — which is exactly
    the crash reported by the user whose code this feature replaces. A test
    helper for this feature should not reproduce that bug.
    """
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        return None
    el = rpr.find(qn("w:highlight"))
    return None if el is None else el.get(qn("w:val"))


def has_highlight(r_el) -> bool:
    return highlight_val(r_el) is not None


def rpr_find(r_el, tag: str):
    """``w:rPr`` child of *r_el* named *tag*, or ``None`` when either is absent.

    ``tag`` is a ``w:``-prefixed name (``"w:color"``), passed straight to
    ``qn()`` so the call site reads like the XML it is asserting on.
    """
    rpr = r_el.find(qn("w:rPr"))
    if rpr is None:
        return None
    return rpr.find(qn(tag))


def add_szcs(run, half_points: int) -> None:
    """Give *run* a ``w:szCs`` after its ``w:sz``.

    python-docx's ``font.size`` writes ``w:sz`` only — no ``w:szCs`` — so a
    fixture that needs the documented ``szCs`` neighbour has to add it. This is
    the same insertion the production code performs, done here so the test can
    assert on a run that genuinely has both.
    """
    rpr = run._r.get_or_add_rPr()
    el = rpr.makeelement(qn("w:szCs"), {})
    el.set(qn("w:val"), str(half_points))
    rpr.insert_element_before(el, "w:highlight", "w:u", "w:effect", "w:bdr",
                              "w:shd", "w:fitText", "w:vertAlign", "w:rtl",
                              "w:cs", "w:em", "w:lang", "w:eastAsianLayout",
                              "w:specVanish", "w:oMath")


def build(path: Path, *runs, style_paragraph=False) -> Path:
    """One paragraph of explicitly specified runs.

    Each ``runs`` entry is ``(text, kwargs)`` where ``kwargs`` goes straight to
    ``run.font`` (``bold``, ``italic``, ``size``, ``name``, ``subscript``).
    """
    doc = Document()
    para = doc.add_paragraph()
    for text, kw in runs:
        run = para.add_run(text)
        for key, value in kw.items():
            setattr(run.font, key, value)
    doc.save(str(path))
    return path


def build_with_szcs(path: Path, text: str, pt: float = 12.0) -> Path:
    """One paragraph, one run, carrying ``w:sz`` **and** ``w:szCs`` **and** ``w:u``.

    Required by the position test: ``w:highlight``'s legal neighbours are
    ``w:szCs`` on the left and ``w:u`` on the right, and python-docx's
    ``font.size`` only writes ``w:sz``. Without this fixture there is no
    ``szCs`` for the "inserted after szCs" assertion to be about.
    """
    doc = Document()
    run = doc.add_paragraph().add_run(text)
    run.font.size = Pt(pt)
    run.font.underline = True
    add_szcs(run, int(pt * 2))
    doc.save(str(path))
    return path


def payload(*args):
    rc, out, err = run_cli(*args)
    assert rc == 0, err
    return json.loads(out)


def runs_of(path: Path, index: int = 0) -> list:
    """``--runs`` record for one paragraph, via the CLI (not via python-docx)."""
    return payload(READ, path, "--runs")["paragraphs"][index]["runs"]


def reloaded(path: Path):
    return Document(str(path))


# ---------------------------------------------------------------------------
# 1. rPr 位置：本包最容易翻车的地方
# ---------------------------------------------------------------------------

def test_highlight_is_inserted_between_szcs_and_u(tmp_path):
    """★ `w:highlight` 必须落在 `w:szCs` 之后、`w:u` 之前 —— 不是「在 rPr 里」。

    这条是方法论 31 的回归钉子：`append` 到末尾同样能得到合法 XML 与
    `ok:true`，只有顺序断言能分辨。fixture 预置 sz/szCs/u 三件（python-docx
    只写 sz，szCs 由 fixture 补），否则没有可比的邻居，位置错也测不出来。
    """
    src = build_with_szcs(tmp_path / "a.docx", "重点", pt=12)
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "重点",
            "--highlight", "darkMagenta", "-o", out)

    tags = rpr_tags(reloaded(out).paragraphs[0].runs[0]._r)
    assert "highlight" in tags, tags
    hi = tags.index("highlight")
    assert "szCs" in tags and "u" in tags, \
        f"fixture 没写出 szCs/u，这条用例失去邻居可比：{tags}"
    assert tags.index("szCs") < hi, f"w:highlight 必须排在 w:szCs 之后：{tags}"
    assert hi < tags.index("u"), f"w:highlight 必须排在 w:u 之前：{tags}"
    # 紧邻关系也钉住：中间冒出别的兄弟说明插入点算错了。
    assert tags[hi - 1] == "szCs" and tags[hi + 1] == "u", \
        f"w:highlight 的左右邻居应是 szCs/u：{tags}"


def test_highlight_position_survives_a_run_without_szcs(tmp_path):
    """没有 sz/u 邻居时也得排在最前面能放的位置，而不是末尾。

    `insert_element_before` 找不到后继元素会 `append`（python-docx 的实现），
    所以这里用 ``color`` 当后继锚：`w:color` 在 `w:highlight` 之前，两者都存在时
    highlight 必须排在 color 后面。
    """
    src = build(tmp_path / "b.docx", ("锚点", {}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "锚点",
            "--color", "FF0000", "--highlight", "yellow", "-o", out)

    tags = rpr_tags(reloaded(out).paragraphs[0].runs[0]._r)
    assert "color" in tags and "highlight" in tags, tags
    assert tags.index("color") < tags.index("highlight"), \
        f"w:color 在 CT_RPr 里先于 w:highlight：{tags}"


def test_repeated_changes_do_not_duplicate_the_element(tmp_path):
    """跑两次不能叠出两个 `w:highlight`（xsd:sequence 每个子元素至多一次）。"""
    src = build(tmp_path / "c.docx", ("重复", {"size": Pt(10)}))
    once = tmp_path / "once.docx"
    twice = tmp_path / "twice.docx"
    payload(EDIT, "char-format", src, "--find", "重复",
            "--highlight", "green", "-o", once)
    payload(EDIT, "char-format", once, "--find", "重复",
            "--highlight", "green", "-o", twice)

    tags = rpr_tags(reloaded(twice).paragraphs[0].runs[0]._r)
    assert tags.count("highlight") == 1, tags
    assert tags.index("sz") < tags.index("highlight"), tags


# ---------------------------------------------------------------------------
# 2. darkMagenta：python-docx 枚举缺的那个值
# ---------------------------------------------------------------------------

def test_dark_magenta_can_be_written(tmp_path):
    """★ 本次真实需求：`w:val="darkMagenta"` 必须能写进去并读回来。

    `WD_COLOR_INDEX` 没有 `DARK_MAGENTA`，`font.highlight_color` 也拒绝这个
    字符串，所以这条同时钉住「实现没退回 python-docx 枚举」。
    """
    src = build(tmp_path / "d.docx", ("事故现场", {}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "事故现场",
                  "--highlight", "darkMagenta", "-o", out)
    assert res["highlight"] == "darkMagenta"

    el = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:highlight")
    assert el is not None
    assert el.get(qn("w:val")) == "darkMagenta"


@pytest.mark.parametrize("spelling", [
    "darkMagenta", "DARK_MAGENTA", "dark_magenta", "DarkMagenta",
    "DARKMAGENTA", "dark-magenta",
])
def test_all_spellings_normalize_to_the_schema_value(tmp_path, spelling):
    """schema 原生拼写 / python-docx 风格 / 小写下划线，都归一化成同一值。"""
    src = build(tmp_path / "e.docx", ("写法", {}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "写法",
                  "--highlight", spelling, "-o", out)
    assert res["highlight"] == "darkMagenta"
    hi = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:highlight")
    assert hi.get(qn("w:val")) == "darkMagenta", spelling


def test_python_docx_enum_alias_is_accepted(tmp_path):
    """`VIOLET` 是 python-docx 给 darkMagenta 起的名字，也收。

    用户从 `docx_read --runs` 拿到的是 schema 拼写；但若他从旧代码里带来
    `WD_COLOR_INDEX.VIOLET`，不该因为名字不同就失败。
    """
    assert WD_COLOR_INDEX.VIOLET.xml_value == "darkMagenta"
    src = build(tmp_path / "f.docx", ("枚举别名", {}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "枚举别名",
                  "--highlight", "VIOLET", "-o", out)
    assert res["highlight"] == "darkMagenta"


def test_every_legal_st_highlight_value_is_accepted(tmp_path):
    """`ST_HighlightColor` 的 17 个值逐个能写 —— 少一个就是一个静默失败面。"""
    sys.path.insert(0, str(SCRIPTS))
    import docx_edit as E

    expected = {
        "black", "blue", "cyan", "darkBlue", "darkCyan", "darkGray",
        "darkGreen", "darkMagenta", "darkRed", "darkYellow", "green",
        "lightGray", "magenta", "none", "red", "white", "yellow",
    }
    assert set(E.HIGHLIGHT_VALUES) == expected

    src = build(tmp_path / "g.docx", ("全部取值", {}))
    for value in sorted(expected - {"none"}):
        out = tmp_path / f"{value}.docx"
        payload(EDIT, "char-format", src, "--find", "全部取值",
                "--highlight", value, "-o", out)
        el = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:highlight")
        assert el is not None, value
        assert el.get(qn("w:val")) == value, value


# ---------------------------------------------------------------------------
# 3. 非法取值：exit 2 + 单行 error:，且不写盘
# ---------------------------------------------------------------------------

@pytest.mark.parametrize("bad", [
    "darkMagento",          # 拼错一个字母，Word 不显示也不报错
    "dark-magenta!",        # 归一化后仍不合法
    "purple",               # 看着像颜色，不是 ST_HighlightColor
    "FF0000",               # 把 --color 的值喂给 --highlight
    "",                     # 空串
])
def test_illegal_highlight_exits_2_with_one_error_line(tmp_path, bad):
    src = build(tmp_path / "h.docx", ("锚", {}))
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "char-format", src, "--find", "锚",
                          "--highlight", bad, "-o", out)
    assert rc == 2, (rc, so, err)
    line = one_error_line(err)
    assert "--highlight" in line and "legal:" in line
    assert so.strip() == "", "拒绝路径不能伪造成功 JSON"
    assert not out.exists(), "拒绝调用不能写盘"


@pytest.mark.parametrize("bad", ["red", "FF00", "GGGGGG", "0xFF0000", "1234567"])
def test_illegal_color_exits_2_with_one_error_line(tmp_path, bad):
    src = build(tmp_path / "i.docx", ("锚", {}))
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "char-format", src, "--find", "锚",
                          "--color", bad, "-o", out)
    # "red" 是 3 位/单词色名，同样拒绝：Word 存的是 RRGGBB，不接受缩写。
    assert rc == 2, (rc, so, err)
    assert "--color" in one_error_line(err)
    assert so.strip() == ""
    assert not out.exists()


def test_no_format_requested_is_refused(tmp_path):
    """一个格式都没给就是空操作；静默成功会让调用方以为改过了。"""
    src = build(tmp_path / "j.docx", ("锚", {}))
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "char-format", src, "--find", "锚", "-o", out)
    assert rc == 2
    assert "at least one format" in one_error_line(err)
    assert so.strip() == ""
    assert not out.exists()


# ---------------------------------------------------------------------------
# 4. 找不到锚点：明确失败，不是 exit 0
# ---------------------------------------------------------------------------

def test_missing_anchor_fails_loudly(tmp_path):
    """★ 找不到锚点必须失败。exit 0 + occurrences:0 是「改了但没变」的伪装。"""
    src = build(tmp_path / "k.docx", ("这里没有那三个字", {}))
    out = tmp_path / "out.docx"
    rc, so, err = run_cli(EDIT, "char-format", src, "--find", "不存在的锚点",
                          "--highlight", "yellow", "-o", out)
    assert rc == 2, (rc, so, err)
    line = one_error_line(err)
    assert "no run in the document carries this text" in line
    assert so.strip() == ""
    assert not out.exists(), "锚点找不到时不能写出一个文件"

    # D-030 判据：拒绝先于任何读写。原文件逐字节不变。
    assert reloaded(src).paragraphs[0].text == "这里没有那三个字"


def test_anchor_that_spans_a_paragraph_break_is_not_found(tmp_path):
    """跨段不算命中：run 不在同一段落里，就没有「包含该文字的那些 run」。"""
    doc = Document()
    doc.add_paragraph("前半")
    doc.add_paragraph("后半")
    src = tmp_path / "l.docx"
    doc.save(str(src))

    rc, _, err = run_cli(EDIT, "char-format", src, "--find", "前半后半",
                         "--highlight", "yellow", "-o", tmp_path / "o.docx")
    assert rc == 2
    assert "no run in the document" in one_error_line(err)


# ---------------------------------------------------------------------------
# 5. 只改锚点里的字：其余 run / 其余字符的格式一律不动
# ---------------------------------------------------------------------------

def test_only_the_anchor_run_is_formatted(tmp_path):
    """整段一起改是最容易犯的错：只有含锚点的 run 变，旁边的不变。"""
    src = build(tmp_path / "m.docx",
                ("前缀", {"size": Pt(12)}),
                ("锚点", {"size": Pt(12)}),
                ("后缀", {"size": Pt(12)}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "锚点",
            "--highlight", "darkMagenta", "-o", out)

    runs = reloaded(out).paragraphs[0].runs
    assert [r.text for r in runs] == ["前缀", "锚点", "后缀"]
    assert [has_highlight(r._r)
            for r in runs] == [False, True, False]


def test_anchor_inside_a_run_splits_it_and_keeps_surrounding_format(tmp_path):
    """锚点只占 run 的一部分时要切开：前后的字保住原格式，只有中间变色。

    这条是「不能整段一起改」的严格版 —— 连同一个 run 里的邻居都不能被牵连。
    """
    src = build(tmp_path / "n.docx",
                ("甲见附录乙", {"size": Pt(14), "bold": True, "name": "KaiTi"}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "见附录",
                  "--highlight", "darkMagenta", "-o", out)
    assert res["runs_changed"] == 1

    runs = reloaded(out).paragraphs[0].runs
    assert [r.text for r in runs] == ["甲", "见附录", "乙"]
    # 三段都是原 run 的深拷贝：bold/size/font 全在
    for run in runs:
        assert run.font.bold is True
        assert run.font.size == Pt(14)
        assert run.font.name == "KaiTi"
    # 只有中间那段带高亮
    flags = []
    for run in runs:
        rpr = run._r.find(qn("w:rPr"))
        flags.append(rpr.find(qn("w:highlight")) is not None)
    assert flags == [False, True, False]


def test_anchor_spanning_three_runs_highlights_all_of_them(tmp_path):
    """锚点跨 run 时，覆盖到的每个 run 都改；没覆盖到的一个都不改。"""
    src = build(tmp_path / "o.docx",
                ("头", {}),
                ("见附", {"size": Pt(12)}),
                ("录一", {"size": Pt(12)}),
                ("尾", {}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "见附录一",
                  "--highlight", "yellow", "-o", out)
    assert res["occurrences"] == 1

    runs = reloaded(out).paragraphs[0].runs
    assert "".join(r.text for r in runs) == "头见附录一尾"
    hi = [has_highlight(r._r)
          for r in runs]
    assert hi[0] is False and hi[-1] is False
    assert sum(hi) >= 2, f"跨 run 的锚点应命中两个以上 run：{hi}"


def test_multiple_occurrences_in_one_paragraph_are_all_changed(tmp_path):
    src = build(tmp_path / "p.docx",
                ("重点", {"size": Pt(12)}),
                ("中间", {"size": Pt(12)}),
                ("重点", {"size": Pt(12)}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "重点",
                  "--highlight", "red", "-o", out)
    assert res["occurrences"] == 2
    runs = reloaded(out).paragraphs[0].runs
    hi = [has_highlight(r._r)
          for r in runs]
    assert hi == [True, False, True]


def test_repeated_anchor_inside_one_run_is_split_twice(tmp_path):
    """同一个 run 里出现两次也要切三刀，而不是「取第一处」。"""
    src = build(tmp_path / "q.docx", ("甲乙甲", {"size": Pt(11)}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "甲",
                  "--highlight", "green", "-o", out)
    assert res["occurrences"] == 2
    runs = reloaded(out).paragraphs[0].runs
    assert [r.text for r in runs] == ["甲", "乙", "甲"]
    hi = [has_highlight(r._r)
          for r in runs]
    assert hi == [True, False, True]


# ---------------------------------------------------------------------------
# 6. 其它格式不被破坏
# ---------------------------------------------------------------------------

def test_other_run_format_is_preserved(tmp_path):
    """改高亮前后，run 的 bold / size / font / italic / vertAlign 一个不少。"""
    src = build(tmp_path / "r.docx",
                ("保留", {"bold": True, "italic": True, "size": Pt(16),
                          "name": "KaiTi"}))
    before = runs_of(src)[0]
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "保留",
            "--highlight", "darkMagenta", "-o", out)
    after = runs_of(out)[0]

    for key in ("bold", "italic", "size_pt", "font", "vert_align"):
        assert after[key] == before[key], f"{key} 被改了：{before} -> {after}"


def test_subscript_survives_highlight(tmp_path):
    """vertAlign 是最容易被切 run 逻辑弄丢的格式；下标字加高亮后仍是下标。"""
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("F")
    run.font.subscript = True
    run.font.size = Pt(14)
    src = tmp_path / "s.docx"
    doc.save(str(src))

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "F",
            "--highlight", "yellow", "-o", out)

    got = runs_of(out)[0]
    assert got["text"] == "F"
    assert got["vert_align"] == "subscript"
    assert got["size_pt"] == 14.0
    assert got["highlight"] == "yellow"


def test_paragraph_level_format_is_untouched(tmp_path):
    """段落对齐不在 run 里，改 run 不该碰到它。"""
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    doc = Document()
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.add_run("居中段")
    src = tmp_path / "t.docx"
    doc.save(str(src))

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "居中段",
            "--highlight", "cyan", "-o", out)
    assert reloaded(out).paragraphs[0].alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_runs_outside_the_anchor_keep_their_own_highlight(tmp_path):
    """只改锚点意味着「没点名的 run 一个字都不许动」，已有高亮也一样。"""
    src = build(tmp_path / "u.docx",
                ("已有高亮", {"size": Pt(12)}),
                ("新锚点", {"size": Pt(12)}))
    # 先给第一段第一个 run 上绿色
    pre = tmp_path / "pre.docx"
    payload(EDIT, "char-format", src, "--find", "已有高亮",
            "--highlight", "green", "-o", pre)

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", pre, "--find", "新锚点",
            "--highlight", "darkMagenta", "-o", out)

    runs = reloaded(out).paragraphs[0].runs
    assert [r.text for r in runs] == ["已有高亮", "新锚点"]
    vals = [highlight_val(r._r)
            for r in runs]
    assert vals == ["green", "darkMagenta"]


# ---------------------------------------------------------------------------
# 7. --color / --bold / --italic
# ---------------------------------------------------------------------------

def test_color_is_written_as_uppercase_hex(tmp_path):
    src = build(tmp_path / "v.docx", ("变色", {}))
    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "变色",
                  "--color", "#00a0ff", "-o", out)
    assert res["color"] == "00A0FF"
    el = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:color")
    assert el.get(qn("w:val")) == "00A0FF"


def test_color_replaces_a_theme_colour(tmp_path):
    """主题色属性优先于 RGB 双胞胎，所以要整颗换掉而不是改 val。

    先手写一个 ``themeColor="accent1"`` 的 run（模拟 Word 内置样式的产物），
    再确认 ``--color`` 之后 themeColor 消失了 —— 留着它，页面颜色不变。
    """
    doc = Document()
    para = doc.add_paragraph()
    run = para.add_run("主题色")
    rpr = run._r.get_or_add_rPr()
    color = rpr.makeelement(qn("w:color"), {})
    color.set(qn("w:val"), "365F91")
    color.set(qn("w:themeColor"), "accent1")
    rpr.append(color)
    src = tmp_path / "w.docx"
    doc.save(str(src))

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "主题色",
            "--color", "000000", "-o", out)

    el = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:color")
    assert el.get(qn("w:themeColor")) is None, "themeColor 留着就还是主题色"
    assert el.get(qn("w:val")) == "000000"


def test_bold_and_italic_flags(tmp_path):
    src = build(tmp_path / "x.docx", ("加粗斜体", {"size": Pt(12)}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "加粗斜体",
            "--bold", "--italic", "-o", out)
    got = runs_of(out)[0]
    assert got["bold"] is True and got["italic"] is True


def test_no_bold_writes_explicit_off_not_absent(tmp_path):
    """``w:val="0"`` 与「元素不存在」不是一回事：后者继承样式里的粗体。"""
    src = build(tmp_path / "y.docx", ("取消粗体", {"bold": True,
                                                "size": Pt(12)}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "取消粗体",
            "--no-bold", "-o", out)

    el = rpr_find(reloaded(out).paragraphs[0].runs[0]._r, "w:b")
    assert el is not None, "必须留下 w:b，而不是删掉它"
    assert el.get(qn("w:val")) == "0"
    assert runs_of(out)[0]["bold"] is False


def test_no_highlight_removes_the_element(tmp_path):
    """``--highlight none`` = 明确关掉；元素被删掉，而不是写个假值。"""
    src = build(tmp_path / "z.docx", ("先有后无", {"size": Pt(12)}))
    pre = tmp_path / "pre.docx"
    payload(EDIT, "char-format", src, "--find", "先有后无",
            "--highlight", "yellow", "-o", pre)
    assert runs_of(pre)[0]["highlight"] == "yellow"

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", pre, "--find", "先有后无",
            "--highlight", "none", "-o", out)
    assert "highlight" not in rpr_tags(reloaded(out).paragraphs[0].runs[0]._r)
    assert runs_of(out)[0]["highlight"] is None


# ---------------------------------------------------------------------------
# 8. --runs 能读回刚设的值（docx_read 新增字段）
# ---------------------------------------------------------------------------

def test_runs_reports_highlight_and_color(tmp_path):
    """★ 设了要能读出来，否则用户无法验证 —— 这是 docx_read 补字段的理由。"""
    doc = Document()
    para = doc.add_paragraph()
    plain = para.add_run("普通")
    marked = para.add_run("标记")
    rpr = marked._r.get_or_add_rPr()
    hi = rpr.makeelement(qn("w:highlight"), {})
    hi.set(qn("w:val"), "darkMagenta")
    rpr.append(hi)
    color = rpr.makeelement(qn("w:color"), {})
    color.set(qn("w:val"), "ff0000")
    rpr.append(color)
    src = tmp_path / "aa.docx"
    doc.save(str(src))

    runs = runs_of(src)
    assert runs[0]["highlight"] is None
    assert runs[0]["color"] is None
    assert runs[1]["highlight"] == "darkMagenta"
    assert runs[1]["color"] == "FF0000", "颜色统一大写，方便与 --color 比对"
    assert plain.text == "普通"


def test_runs_reports_schema_spelling_not_enum_name(tmp_path):
    """读回的是 ``darkMagenta``，不是 python-docx 的 ``VIOLET``。

    否则「读到的值」喂回 `--highlight` 还得再对照一遍表，且枚举里根本没有
    ``DARK_MAGENTA``，用户无从得知 schema 拼写。
    """
    src = build(tmp_path / "ab.docx", ("枚举名", {}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "枚举名",
            "--highlight", "darkMagenta", "-o", out)
    got = runs_of(out)[0]["highlight"]
    assert got == "darkMagenta"
    assert got.upper() != "VIOLET"


def test_runs_reports_theme_colour_distinctly(tmp_path):
    """主题色要报成 ``theme:accent1``，不能报成某个 hex —— 它有自己的一层优先级。"""
    doc = Document()
    run = doc.add_paragraph().add_run("主题")
    color = run._r.get_or_add_rPr().makeelement(qn("w:color"), {})
    color.set(qn("w:val"), "365F91")
    color.set(qn("w:themeColor"), "accent1")
    run._r.find(qn("w:rPr")).append(color)
    src = tmp_path / "ac.docx"
    doc.save(str(src))

    assert runs_of(src)[0]["color"] == "theme:accent1"


def test_runs_row_has_round_trip_shape(tmp_path):
    """`--runs` 只能加键：既有键一个不少，且新键可被 char-format 直接消费。"""
    src = build(tmp_path / "ad.docx", ("回路", {"bold": True,
                                              "size": Pt(12)}))
    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "回路",
            "--highlight", "darkMagenta", "-o", out)

    run = runs_of(out)[0]
    for key in ("text", "vert_align", "bold", "italic", "size_pt", "font",
                "is_field", "is_drawing", "highlight", "color"):
        assert key in run, f"--runs 丢了既有键 {key}：{sorted(run)}"

    # 读回来的值能原样喂回去（归一化是幂等的）
    again = tmp_path / "again.docx"
    payload(EDIT, "char-format", out, "--find", "回路",
            "--highlight", run["highlight"], "--color", "00A0FF",
            "-o", again)
    assert runs_of(again)[0]["highlight"] == "darkMagenta"
    assert runs_of(again)[0]["color"] == "00A0FF"


def test_blocks_contract_is_unchanged(tmp_path):
    """新字段只进 `--runs`：`--blocks`（markdown 往返路径）一字不变。"""
    src = build(tmp_path / "ae.docx", ("块", {}))
    flat = json.dumps(payload(READ, src, "--blocks"))
    assert "highlight" not in flat
    assert "mixed_format" not in flat


# ---------------------------------------------------------------------------
# 9. 容器覆盖：表格 / 页眉页脚；以及明确的射程外
# ---------------------------------------------------------------------------

def test_table_cell_runs_are_reachable(tmp_path):
    """表格里的字也算「文档里的文字」，不能只在 body 段落里找。"""
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).paragraphs[0].add_run("格内锚点")
    src = tmp_path / "af.docx"
    doc.save(str(src))

    out = tmp_path / "out.docx"
    res = payload(EDIT, "char-format", src, "--find", "格内锚点",
                  "--highlight", "darkMagenta", "-o", out)
    assert res["occurrences"] == 1
    run = reloaded(out).tables[0].cell(0, 0).paragraphs[0].runs[0]
    assert highlight_val(run._r) == "darkMagenta"


def test_header_and_footer_are_reachable(tmp_path):
    """页眉页脚与 `replace` 同射程：它们也是用户会在意的可见文字。"""
    doc = Document()
    doc.add_paragraph("正文")
    doc.sections[0].header.paragraphs[0].add_run("页眉锚点")
    doc.sections[0].footer.paragraphs[0].add_run("页脚锚点")
    src = tmp_path / "ag.docx"
    doc.save(str(src))

    out = tmp_path / "out.docx"
    payload(EDIT, "char-format", src, "--find", "页眉锚点",
            "--highlight", "yellow", "-o", out)
    payload(EDIT, "char-format", out, "--find", "页脚锚点",
            "--highlight", "cyan", "-o", out)

    got = reloaded(out)
    hp = got.sections[0].header.paragraphs[0].runs[0]._r
    fp = got.sections[0].footer.paragraphs[0].runs[0]._r
    assert highlight_val(hp) == "yellow"
    assert highlight_val(fp) == "cyan"


def test_refusal_leaves_the_input_file_byte_identical(tmp_path):
    """D-030：拒绝先于任何读写。整颗文件的字节都不许动。"""
    src = build(tmp_path / "ah.docx", ("原样", {"size": Pt(12)}))
    before = src.read_bytes()
    rc, _, _ = run_cli(EDIT, "char-format", src, "--find", "原样",
                       "--highlight", "notacolour")
    assert rc == 2
    assert src.read_bytes() == before


def test_in_place_run_is_allowed(tmp_path):
    """不给 `-o` 就是就地改；这是既有 8 个子命令的一致行为。"""
    src = build(tmp_path / "ai.docx", ("就地", {"size": Pt(12)}))
    res = payload(EDIT, "char-format", src, "--find", "就地",
                  "--highlight", "magenta")
    assert res["output"] == str(src)
    assert runs_of(src)[0]["highlight"] == "magenta"
