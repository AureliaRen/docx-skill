# -*- coding: utf-8 -*-
"""``docx_validate`` 的受限 rPr 枚举值检查（P20-A）。

## 本文件存在的理由

P19-C 用 Word COM 逐个 ``Documents.Open`` 实测出一批**「校验器说健康、Word 说
文件可能已经损坏」**的文件：`<w:highlight w:val="DARK_MAGENTA">`、
`w:val="purple"`、`w:val="FF0000"`、`w:val="Magenta"`、`w:val="darkMagento"`
共 6 个，全部被 `docx_validate` 报 `ok:true`。这不是格式丢失，是**文档彻底打
不开**——已知盲区里后果最严重的一条。

所以本文件钉三件事：

1. **该抓的抓到**：`w:highlight` / `w:vertAlign` / `w:effect` 的非法值 →
   ``invalid-run-property``（error）+ exit 1；
2. **不该响的不响**（这条比上面important）：17 个合法 highlight 值、``auto``、
   缺 ``w:val`` 的 ``w:u`` / ``w:color``、没有 rPr 的普通文档 → ``ok:true``。
   这个报告被 `docx_edit` / `docx_template` / `docx_mathify` 下游消费，
   误报会让**好文档被拒收**；
3. **severity 按实测分工**：`w:u` / `w:color` 的非法值是
   ``nonstandard-run-property``（**warning**，exit 0）—— 实测 Word 照常打开，
   只是重新解释。凭 schema 判它 error 就是误报。

P21 又补了两条**方向相反**的用例，钉住「严重度由 Word 实测决定、不由 schema
决定」这条总原则：``vertAlign="normal"``（schema 没有、Word 能开）必须放行，
``effect="ants"``（schema 有、Word 打不开）必须是 error。只修一条，原则就是
半截的。

## 为什么**不**挂 ``needs_omml`` / ``needs_word``

纯 XML 层：只用 ``zipfile`` + ``lxml`` + ``python-docx`` 造包与读包，不碰
Office COM / LibreOffice / ``MML2OMML.XSL``。CI runner 没有 Office，必须照常
pass。Word 实测结果作为**依据**写在注释里，不在这里复跑。

Run:  python -m pytest tests/test_validate_enums.py -q
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

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SMOKE = SKILL.parent / "_smoke"
BASE = SMOKE / "in.docx"

#: 本检查的两条 code。error 那条让文件打不开；warning 那条 Word 能打开。
ERROR_CODE = "invalid-run-property"
WARN_CODE = "nonstandard-run-property"


# ---- helpers -----------------------------------------------------------------

def run_validate(path) -> subprocess.CompletedProcess:
    env = dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
    return subprocess.run(
        [sys.executable, str(SCRIPT), str(path)],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False)


def report_of(proc) -> dict:
    """stdout 必须整份是 JSON —— 校验器不许拿报告换 traceback。"""
    return json.loads(proc.stdout)


def codes(rep: dict) -> list:
    return [i["code"] for i in rep["issues"]]


def enum_issues(rep: dict) -> list:
    return [i for i in rep["issues"] if i["code"] in (ERROR_CODE, WARN_CODE)]


def patch_run_properties(src: Path, dst: Path, snippet: str) -> Path:
    """往 ``src`` 的**第一个 run** 上挂一段 ``rPr`` 子元素，产出 ``dst``。

    其余部件逐字节原样复制（``writestr`` 用同一个 ``ZipInfo``），所以包里唯一
    的差异就是 ``snippet`` —— 测试失败时不会混进别的变量。
    """
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/document.xml":
                text = data.decode("utf-8")
                at = text.find("<w:r>")
                if at < 0:                      # 带属性的 run 开标签
                    at = text.find("<w:r ")
                    at += 5
                else:
                    at += 5
                text = (text[:at]
                        + '<w:rPr><w:rFonts w:ascii="Times New Roman" '
                          'w:hAnsi="Times New Roman"/><w:sz w:val="24"/>'
                        + snippet + "</w:rPr>"
                        + text[at:])
                data = text.encode("utf-8")
            zout.writestr(item, data)
    return dst


@pytest.fixture(scope="module")
def base_docx() -> Path:
    if not BASE.exists():               # _smoke 只在开发树里有，CI 上跳过
        pytest.skip("_smoke/in.docx not shipped with the repo")
    return BASE


def build(base: Path, tmp_path: Path, snippet: str, name: str) -> Path:
    return patch_run_properties(base, tmp_path / name, snippet)


# ---- 该抓的：rPr 里非法值让 Word 拒绝打开 → error + exit 1 ---------------------

#: 每条都是 P19-C 用 Word 16.0 COM ``Documents.Open`` 实测过「文件可能已经损坏」
#: (com_error 25272) 的值，括号里是实测结果。
FATAL = [
    ('highlight', 'DARK_MAGENTA'),      # 大小写：OOXML 只认首字母小写
    ('highlight', 'Magenta'),
    ('highlight', 'purple'),            # 不在 ST_HighlightColor 里
    ('highlight', 'FF0000'),            # 十六进制不是颜色名
    ('highlight', 'darkMagento'),       # 拼写错
    ('highlight', 'inherit'),           # 命令行语义词，不是 schema 值
    ('vertAlign', 'Normal'),            # 大小写错（小写 normal 能打开，见 README_OF_FIX）
    ('vertAlign', 'SUPERSCRIPT'),
    ('vertAlign', 'Baseline'),
    ('vertAlign', 'bogus'),
    ('effect', 'blink'),                # 合法值是 blinkBackground
    ('effect', 'ants'),                 # P21：schema 合法，Word 实测拒绝打开
]


@pytest.mark.parametrize("tag,value", FATAL)
def test_illegal_value_is_an_error(base_docx, tmp_path, tag, value):
    path = build(base_docx, tmp_path, f'<w:{tag} w:val="{value}"/>',
                 f"{tag}_{value}.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 1, (tag, value, proc.stdout, proc.stderr)
    assert rep["ok"] is False
    got = enum_issues(rep)
    assert len(got) == 1, got
    assert got[0]["code"] == ERROR_CODE
    assert got[0]["severity"] == "error"
    # detail 必须点出是哪个元素、哪个值 —— 只说「有个 run 属性不对」没法修
    assert f"<w:{tag}" in got[0]["detail"]
    assert value in got[0]["detail"]


@pytest.mark.parametrize("tag", ['highlight', 'vertAlign', 'effect'])
def test_missing_val_is_an_error(base_docx, tmp_path, tag):
    """``<w:highlight/>`` 没有 w:val：schema 里 w:val 是 required，Word 拒绝打开。

    实测（Word COM）：``<w:highlight/>`` / ``<w:vertAlign/>`` / ``<w:effect/>``
    三者全部 com_error 25272。
    """
    path = build(base_docx, tmp_path, f"<w:{tag}/>", f"{tag}_noval.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 1
    assert rep["ok"] is False
    got = enum_issues(rep)
    assert [i["code"] for i in got] == [ERROR_CODE]
    assert "no w:val" in got[0]["detail"]


# ---- P21 回归：severity 由 Word 实测决定，不由 schema 决定 ---------------------
#
# 这一节专门钉住两个「schema 与 Word 不一致」的值。两条都是审查挖出来的真
# 缺陷，且方向相反 —— 一条是误报（把能打开的文档判成坏的），一条是漏报（把
# 打不开的文档判成健康的）。它们必须成对存在：只修一条就会把原则做成半截。

def test_vert_align_normal_is_reported_as_clean(base_docx, tmp_path):
    """``<w:vertAlign w:val="normal">`` → ``ok:true`` + exit 0（P21 修复 1）。

    修前实测（本用例在原代码上跑）::

        {"ok": false, "issues": [{"severity": "error",
          "code": "invalid-run-property",
          "detail": "... <w:vertAlign w:val='normal'> is not a legal value;
                     legal: baseline, subscript, superscript"}]}
        exit=1

    而 Word 16.0 COM 实测：``Documents.Open`` OK，``Superscript=0``（正常）、
    ``Subscript=0``，存盘后该元素被丢弃。也就是说这个文件**能正常打开、效果
    等同 baseline**，校验器却让下游拒收它。

    更难看的是同一个文件的注释第 141 行附近本来就写着「opens」，代码里却把
    它设成 error —— 注释与代码互斥。本用例把结论钉死：判 error 就是缺陷。
    """
    path = build(base_docx, tmp_path, '<w:vertAlign w:val="normal"/>',
                 "va_normal.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 0, proc.stdout
    assert rep["ok"] is True
    # 连 warning 都不该有：Word 存盘时自己会丢掉这个元素，没有任何内容损失。
    assert rep["issues"] == [], rep["issues"]
    assert enum_issues(rep) == []


def test_effect_ants_is_an_error(base_docx, tmp_path):
    """``<w:effect w:val="ants">`` → ``invalid-run-property`` + exit 1（P21 修复 2）。

    ``ants`` **在** ``ST_TextEffect`` 里（五个合法值之一），但 Word 16.0 实测
    拒绝打开（com_error -2147352567）。修前的代码把 ``ants`` 放进合法集合，
    于是返回 ``{"ok": true}``，并在注释里解释「故意漏掉」—— 审查的批评对：
    故意漏掉应当表现为**不在合法集合**，而不是放进集合再解释。

    判据统一为「Word 实测」，所以这里断言它和 ``blink`` 同档。
    """
    path = build(base_docx, tmp_path, '<w:effect w:val="ants"/>',
                 "effect_ants.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 1, proc.stdout
    assert rep["ok"] is False
    got = enum_issues(rep)
    assert [i["code"] for i in got] == [ERROR_CODE]
    assert "ants" in got[0]["detail"]


def test_effect_legal_set_is_the_word_set_not_the_schema_set(base_docx):
    """合法 ``effect`` 集合 = schema 五值 **减 ants**，且不含 ``blink``。

    直接读规则表，而不是靠上面那些用例的副作用推：这张表是「Word 实测结论」
    的唯一落点，把 ``ants`` 加回去必须让本用例变红。
    """
    from docx_validate import _EFFECT_VALUES
    assert set(_EFFECT_VALUES) == {"blinkBackground", "lights", "sparkle",
                                   "none"}
    assert "ants" not in _EFFECT_VALUES
    assert "blink" not in _EFFECT_VALUES        # 那是 blinkBackground 的错拼


def test_vert_align_legal_set_takes_both_sides_of_the_table(base_docx):
    """合法 ``vertAlign`` 集合取自 ``VERT_ALIGN`` 的键**和**值。

    只取 values 会丢掉 ``normal``（它映射到 ``None``），那正是修复 1 的成因；
    只取 keys 会丢掉能打开的 ``baseline`` / ``superscript`` / ``subscript``。
    两侧都要，且 ``None`` 不能混进来。
    """
    from docx_crossref import VERT_ALIGN
    from docx_validate import _legal_vert_align
    legal = _legal_vert_align()
    assert "normal" in legal
    assert set(VERT_ALIGN.values()) - {None} <= set(legal)
    assert None not in legal
    assert "" not in legal


def test_word_opens_but_schema_forbids_is_never_an_error(base_docx, tmp_path):
    """把「Word 能打开」这一条性质做成表驱动断言，防止将来再按 schema 判档。

    ``w:u`` / ``w:color`` 全部落在 warning，``vertAlign=normal`` 落在零命中
    —— 三种「schema 说不行、Word 说行」的值都不许把 ``ok`` 翻成 false。
    """
    cases = [
        ('<w:u w:val="SINGLE"/>', [WARN_CODE]),
        ('<w:color w:val="#FF0000"/>', [WARN_CODE]),
        ('<w:vertAlign w:val="normal"/>', []),
    ]
    for snippet, expected in cases:
        path = build(base_docx, tmp_path, snippet, "wbx.docx")
        proc = run_validate(path)
        rep = report_of(proc)
        assert proc.returncode == 0, (snippet, proc.stdout)
        assert rep["ok"] is True, (snippet, rep)
        assert [i["code"] for i in enum_issues(rep)] == expected, (snippet, rep)


# ---- 不该响的：合法值必须零命中 ----------------------------------------------

#: ``ST_HighlightColor`` 全 17 个合法值，逐个用 Word COM 打开实测 **全部 OK**。
LEGAL_HIGHLIGHT = (
    "black", "blue", "cyan", "darkBlue", "darkCyan", "darkGray", "darkGreen",
    "darkMagenta", "darkRed", "darkYellow", "green", "lightGray", "magenta",
    "none", "red", "white", "yellow",
)


@pytest.mark.parametrize("value", LEGAL_HIGHLIGHT)
def test_legal_highlight_is_clean(base_docx, tmp_path, value):
    """尤其 ``darkMagenta``：它就是被要求支持的那个值，绝不能报。"""
    path = build(base_docx, tmp_path, f'<w:highlight w:val="{value}"/>',
                 f"ok_hl_{value}.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 0, (value, proc.stdout)
    assert rep["ok"] is True
    assert enum_issues(rep) == []


@pytest.mark.parametrize("value", ["baseline", "normal", "subscript",
                                  "superscript"])
def test_legal_vert_align_is_clean(base_docx, tmp_path, value):
    """``normal`` 必须在里面（P21）。

    ``ST_VerticalAlignRun`` 没有 ``normal``，但 Word 16.0 实测**照常打开**，
    效果等同 baseline（COM 读到 ``Superscript=0 Subscript=0``），存盘时把元素
    丢掉。它不是「Word 容忍的乱写」，而是 ``docx_crossref.VERT_ALIGN`` 里
    ``normal -> None`` 的写侧拼法、也是 ``--vert-align`` 的默认值 —— 把它判
    error 等于拒收本仓自己写出来的文档。所以它进合法集，依据是 Word 实测，
    不是 schema。反例见 ``FATAL`` 里的 ``Normal``：大小写换一下就真的打不开。
    """
    path = build(base_docx, tmp_path, f'<w:vertAlign w:val="{value}"/>',
                 f"ok_va_{value}.docx")
    proc = run_validate(path)
    assert proc.returncode == 0, (value, proc.stdout)
    assert enum_issues(report_of(proc)) == []


@pytest.mark.parametrize("value", ["blinkBackground", "lights", "sparkle",
                                  "none"])
def test_legal_effect_is_clean(base_docx, tmp_path, value):
    """Word 能打开的 ``ST_TextEffect`` 值一个都不许响。

    ``sparkle`` 打开后 Word 存成 ``none``（重新解释），仍然算「能打开」，所以
    这里也必须是干净的 —— 判据是打开与否，不是存盘后是否原样保留。
    """
    path = build(base_docx, tmp_path, f'<w:effect w:val="{value}"/>',
                 f"ok_effect_{value}.docx")
    assert enum_issues(report_of(run_validate(path))) == []


@pytest.mark.parametrize("value", [
    "single", "none", "double", "dotted", "dash", "dotDash", "dotDotDash",
    "wave", "thick", "words", "dottedHeavy", "dashedHeavy", "dashDotHeavy",
    "dashDotDotHeavy", "wavyHeavy", "dashLong", "wavyDouble", "dashLongHeavy",
])
def test_legal_underline_is_clean(base_docx, tmp_path, value):
    path = build(base_docx, tmp_path, f'<w:u w:val="{value}"/>',
                 f"ok_u_{value}.docx")
    assert enum_issues(report_of(run_validate(path))) == []


@pytest.mark.parametrize("value", ["auto", "FF0000", "000000", "ff00aa"])
def test_legal_color_is_clean(base_docx, tmp_path, value):
    path = build(base_docx, tmp_path, f'<w:color w:val="{value}"/>',
                 f"ok_c_{value}.docx")
    assert enum_issues(report_of(run_validate(path))) == []


@pytest.mark.parametrize("snippet", [
    "<w:u/>",                       # 缺 w:val = single（实测 Word 打开 OK）
    "<w:color/>",                   # 缺 w:val = auto（实测 Word 打开 OK）
    '<w:u w:color="ZZZZZZ"/>',      # u 上的 w:color 是颜色不是下划线类型
    '<w:color w:val="auto" w:themeColor="accent1"/>',
])
def test_absent_or_unrelated_value_is_clean(base_docx, tmp_path, snippet):
    """缺 ``w:val`` 只有那三个 required 的元素才算错；u / color 各有默认值。

    另外钉住「只看元素自己的名字」：``w:u`` 也带 ``w:color`` 属性，那不是
    ``ST_Underline`` 的值，混起来查就会误报。
    """
    path = build(base_docx, tmp_path, snippet, "ok_absent.docx")
    assert enum_issues(report_of(run_validate(path))) == []


def test_plain_document_without_rpr_is_clean(tmp_path):
    """``Document()`` 造的空文档：正文没有 rPr，styles 里的那一堆必须不响。"""
    from docx import Document
    path = tmp_path / "plain.docx"
    Document().save(str(path))

    proc = run_validate(path)
    rep = report_of(proc)
    assert proc.returncode == 0, proc.stdout
    assert rep["ok"] is True
    assert enum_issues(rep) == []


def test_real_smoke_documents_are_clean():
    """本机 ``_smoke`` 真实语料：一个 rPr 命中都不能有。"""
    docs = sorted(SMOKE.glob("*.docx"))
    if not docs:
        pytest.skip("_smoke/*.docx not shipped with the repo")
    for doc in docs:
        rep = report_of(run_validate(doc))
        assert enum_issues(rep) == [], (doc.name, rep["issues"])


# ---- 不该响的（但值得说一句）：Word 能打开 → warning, exit 0 -------------------

@pytest.mark.parametrize("snippet,tag", [
    ('<w:u w:val="SINGLE"/>', "u"),          # 大小写错，Word 重新解释
    ('<w:u w:val="tripleWave"/>', "u"),      # 不在 ST_Underline 里
    ('<w:color w:val="F00"/>', "color"),     # 三位缩写
    ('<w:color w:val="#FF0000"/>', "color"),  # 带 # —— 那是 CSS 写法
    ('<w:color w:val="red"/>', "color"),     # 颜色名不是十六进制
    ('<w:color w:val="ZZZZZZ"/>', "color"),  # 六个字符，但不是十六进制
    ('<w:color w:val=""/>', "color"),        # 空值
])
def test_out_of_schema_but_word_opens_is_a_warning(base_docx, tmp_path,
                                                   snippet, tag):
    """这 7 个都是实测 **Word 照常打开** 的值。

    凭 schema 判它们 error 就是误报 —— 下游会拒收能打开的文档。所以定
    ``warning``：``ok`` 仍是 ``true``，退出码仍是 0，但报告里说清楚。
    """
    path = build(base_docx, tmp_path, snippet, f"warn_{tag}.docx")
    proc = run_validate(path)
    rep = report_of(proc)

    assert proc.returncode == 0, (snippet, proc.stdout)
    assert rep["ok"] is True
    got = enum_issues(rep)
    assert [i["code"] for i in got] == [WARN_CODE]
    assert got[0]["severity"] == "warning"


def test_warning_does_not_pollute_ok(base_docx, tmp_path):
    """warning 只出现在 issues 里，不许把 ``ok`` 翻成 false（P17 的 severity 契约）。"""
    path = build(base_docx, tmp_path, '<w:color w:val="red"/>', "warn_only.docx")
    rep = report_of(run_validate(path))
    assert rep["ok"] is True
    assert any(i["severity"] == "warning" for i in rep["issues"])


# ---- 覆盖范围：不止 document.xml ----------------------------------------------

@pytest.mark.parametrize("part", [
    "word/styles.xml",
    "word/stylesWithEffects.xml",
    "word/numbering.xml",
    "word/settings.xml",
])
def test_illegal_value_in_any_run_property_part_is_caught(base_docx, tmp_path,
                                                          part):
    """非法值写进 styles / numbering / settings 一样让 Word 打不开（实测）。

    只扫 ``document.xml`` 会漏掉**绝大多数**真实情况：语料里 4799 个 rPr 只有
    200 个在 document.xml，其余全在 styles / stylesWithEffects 里。
    """
    src = base_docx
    dst = tmp_path / ("part_" + Path(part).stem + ".docx")
    with zipfile.ZipFile(src) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == part:
                text = data.decode("utf-8")
                at = text.find("<w:rPr>")
                if at >= 0:
                    text = (text[:at + 7] + '<w:highlight w:val="purple"/>'
                            + text[at + 7:])
                else:
                    # ``word/settings.xml`` 本身没有 rPr。挂到**根元素**里 ——
                    # 命名空间声明在根上，插到 XML 声明后面会让 ``w:`` 前缀
                    # 悬空、整份 part 变坏（那测到的就不是枚举检查了）。
                    root_end = text.index(">", text.index("?>") + 2) + 1
                    text = (text[:root_end]
                            + '<w:rPr><w:highlight w:val="purple"/></w:rPr>'
                            + text[root_end:])
                data = text.encode("utf-8")
            zout.writestr(item, data)

    proc = run_validate(dst)
    rep = report_of(proc)
    assert proc.returncode == 1, (part, proc.stdout)
    got = enum_issues(rep)
    assert len(got) == 1
    assert got[0]["code"] == ERROR_CODE
    assert part in got[0]["detail"]


def test_several_bad_values_are_each_reported(base_docx, tmp_path):
    """一个 run 上挂两个非法元素 → 两条 issue，不是一条。"""
    path = build(base_docx, tmp_path,
                 '<w:highlight w:val="purple"/><w:vertAlign w:val="bogus"/>',
                 "two_bad.docx")
    rep = report_of(run_validate(path))
    got = enum_issues(rep)
    assert len(got) == 2
    details = " ".join(i["detail"] for i in got)
    assert "<w:highlight" in details and "<w:vertAlign" in details


# ---- 不崩：坏部件只跳过，不许换掉整份报告 -------------------------------------

def test_unreadable_part_is_skipped_not_fatal(base_docx, tmp_path):
    """``word/settings.xml`` 不是良构 XML 时，枚举检查静默跳过。

    同一个坏部件已经由别处报 ``bad-*-xml`` 或由 python-docx 拒绝；再补一条
    枚举 issue 只是噪音。要紧的是**报告还在**（JSON on stdout，不是栈帧）。
    """
    dst = tmp_path / "broken_settings.docx"
    with zipfile.ZipFile(base_docx) as zin, \
            zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                data = b"<w:settings"
            zout.writestr(item, data)

    proc = run_validate(dst)
    rep = report_of(proc)                # 不许是 traceback
    assert enum_issues(rep) == []
    assert proc.stdout.strip().startswith("{")


def test_report_shape_is_unchanged_for_healthy_doc(base_docx):
    """健康文档的报告形状不变：``{"ok": true, "issues": []}``。"""
    rep = report_of(run_validate(base_docx))
    assert rep == {"ok": True, "issues": []}
