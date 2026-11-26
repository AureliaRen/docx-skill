# -*- coding: utf-8 -*-
r"""定界组空槽：组首一元正负号被 MML2OMML.XSL 误判成槽分隔符。

## 这个 bug 长什么样

`\\left(-\\frac{y}{x}\\right)` 或正文里更常见的 `(-\\frac{y}{x})` 转成 OMML 后是

    <m:d><m:dPr><m:sepChr m:val="−"/></m:dPr><m:e/><m:e>…</m:e></m:d>

第一个 `m:e` 是**空的**，Word 把它画成一个空方块。根因在 XSL 而不在本仓库：
把 latex2mathml 的 MathML 直接喂 `math_omml.xslt()` 就复现，连 skill 代码都不经过。

## 为什么判据不是 `fence`

`\\left( … \\right)` 产的 `<mo>` 带 `fence="true"`，**普通括号不带**（只有
`stretchy="false"`）。所以「首尾都是 fence」的判据只能修到 `\\left` 那一半，
而正文里 `(-\\frac{y}{x})` 这种普通括号写法同样会坏 —— 实测同样 1 个空 `m:e`。
判据因此改成**平衡深度**：首元素是开定界符、末元素是对应的闭定界符、且中间深度
从不归零。归零说明中途闭合过，那是并排的几个组（`(-x)(-y)`）不是一个组。

## 三条反直觉但正确的期望

1. `(-x)` 修完是 **0** 个空槽，但 `\\left(-x+y\\right)` 修完 **sep 仍是 `+`**：
   前者 wrap 把「−x」包成一体，组内只剩一项；后者 wrap 只包「−x」，后面那个 `+`
   留在组内当分隔符 —— 行为不变，这是设计而不是回归。
2. `\\left(\\right)`（空组）**修不好也不该修**：组内 0 个项、没有第二项可包。
   它必须落到 `assert_legal` 的新检查上抛错，而不是被编造内容。
3. `\\sqrt{}` 产生的空 `m:rad/m:e` **不能**被新检查判非法 —— 那条已有专门检查，
   且 `m:e` 在 `m:num` / `m:fName` 里也合法、也常为空。新检查只认 `m:d` 直接子元素。

OMML 可用性判定统一走 `tests/_omml.py`（全仓唯一实现）。
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import math_omml  # noqa: E402
from _omml import needs_omml  # noqa: E402

M = math_omml.M_NS


def _m(tag: str) -> str:
    return "{%s}%s" % (M, tag)


def empty_slots(el) -> int:
    """`m:d` 下空 `m:e` 的个数（只数 `m:d` 直接子元素）。

    刻意**只**数 `m:d`：`m:rad/m:e` 空着是另一回事（见 `test_new_check_ignores_rad`）。
    """
    return sum(1 for d in el.iter(_m("d"))
               for c in d
               if c.tag == _m("e") and len(c) == 0)


def seps(el) -> list[str]:
    """所有 `m:sepChr@val`（文档序）。断言「正常分隔符不被改坏」用这个。"""
    return [s.get(_m("val")) for s in el.iter(_m("sepChr"))]


def convert(latex: str, stats: dict | None = None):
    return math_omml.latex_to_omml_lxml(latex, stats=stats)


# ------------------------------------------------------------------ 1. 修好的
# 每一条都实测过：before 是 1（只有双定界组 `\\left(-\\frac{1}{2}\\right)\\left(+\\frac{a}{b}\\right)` 是 2），
# after 必须是 0。分成两组是因为它们的**成因入口不同** —— 一组走 `\\left…\\right`，
# 一组走普通括号；`fence` 判据只能抓到前一组，这正是本单返工的原因。
FIXED_LEFT_RIGHT = [
    r"\left(-x\right)",
    r"\left(-\frac{y}{x}\right)",
    r"\left(+\frac{a}{b}\right)",
    r"\left(-1\right)^{n}",
    r"\left[-x\right]",
    r"\left|-x\right|",
    r"\left\{-\frac{y}{x}\right\}",
    r"\left(\mp x\right)",
    r"\left(-\mathbf{v}\right)",
    r"\left(-\text{if }x\right)",
]

FIXED_PLAIN = [
    r"(-\frac{y}{x})",
    r"(-x)",
    r"(-1)",
    r"(-x y)",
    r"(-\sin x)",
    r"(-\sqrt{2})",
    r"(-a^{2})",
    r"|-x|",
    r"[-x]",
    r"\{-x\}",
    r"(\mp x)",
]

@needs_omml
@pytest.mark.parametrize("latex", FIXED_LEFT_RIGHT + FIXED_PLAIN)
def test_leading_unary_group_has_no_empty_slot(latex):
    """组首一元正负号的定界组，修完必须是 0 个空槽。

    普通括号那 11 条是本单最容易漏的一片面：它们不带 `fence` 属性，
    只看 `fence` 的判据一条都修不到。
    """
    el = convert(latex)
    assert empty_slots(el) == 0, f"{latex!r} 仍有空槽"


@needs_omml
@pytest.mark.parametrize("latex, wraps", [
    (r"\left(\left(-x\right)\right)", 1),
    (r"\left(-\frac{1}{2}\right)\left(+\frac{a}{b}\right)", 2),
    (r"f(x)=\left(-\sin x\right)+\left(\cos y\right)", 1),
    (r"V'(y)=[k(x-b)+2F]\left(-\frac{y}{x}\right)=0", 1),
])
def test_nested_and_multiple_groups(latex, wraps):
    """嵌套与多组：每一处组首一元号各 wrap 一次，空槽一律 0。

    嵌套那条 before 是 **1** 个空槽（只有内层：外层的首槽装着内层 `m:d`，非空），
    所以它验证的是「内层被 wrap 之后外层仍被正确判定为无需再包」——
    `iter()` 的文档序先给外层，外层 body[0] 是定界符而不是一元号，一条都不命中。
    """
    stats: dict = {}
    el = convert(latex, stats)
    assert empty_slots(el) == 0
    assert stats.get("delim_wraps") == wraps


# ------------------------------------------------------------------ 2. 不许动的
# 这一组修完必须**仍然** 0 空槽。`sepChr` 要分两类说：
#   * `KEEP_UNTOUCHED` —— 一个都不变，与修复前逐字节相同；
#   * `KEEP_WRAPPED`   —— 会从 `''` 变成组内**真正的运算符**（`+` / `,`），
#     即「一个槽装全部、sep 空」→「两个槽、sep 是真运算符」。**结构变了**，
#     渲染等价：Word 画「槽1 + sep + 槽2」，与原来单槽里的同一串字符观感一致。
#
# 两类混着写会带出两个误会：以为 wrap 次数是判据（它不是），以及以为
# `KEEP_WRAPPED` 的 sep 不变（它变了，`want_seps` 填的就是修复后的真值）。
KEEP_UNTOUCHED = [
    (r"\left(a+b\right)", ["+"]),
    (r"\left(a,\ b\right)", [","]),
    (r"\left(\frac{a}{b}\right)", [""]),
    (r"\left(\pm\frac{a}{b}\right)", [""]),
    (r"(\pm x)", [""]),
    (r"(a)(b)", [""]),
    (r"(a+b)(c-d)", [""]),
    (r"(-x)(-y)", [""]),
    (r"(-x)-(-y)", [""]),
    (r"(-x)^{2}", [""]),
    (r"a(-x)b", []),
    (r"y=(-x)+1", []),
]

KEEP_WRAPPED = [
    (r"\left(-x+y\right)", ["+"]),
    (r"(-\frac{a}{b}+1)", ["+"]),
    (r"(-x,\ y)", [","]),
]

@needs_omml
@pytest.mark.parametrize("latex, want_seps", KEEP_UNTOUCHED)
def test_normal_delimiters_untouched(latex, want_seps):
    """正常分隔符：修完 sep 值一个都不变，且一次 wrap 都不该发生。

    `\\left(a+b\\right)` 的 sep 必须是 `+` —— 这是本单最容易改坏的一条：
    只要判据放宽成「组里有运算符就包」，它就会从两个槽塌成一个。
    `(-x)(-y)` 与 `(-x)-(-y)` 钉的是**深度平衡**：首尾看着像一对括号，但中间
    深度归零过，那是并排的两个组，不能包。
    """
    stats: dict = {}
    el = convert(latex, stats)
    assert empty_slots(el) == 0
    assert seps(el) == want_seps, f"{latex!r} 的 sep 被改了"
    assert stats.get("delim_wraps", 0) == 0, f"{latex!r} 不该被 wrap"


@needs_omml
@pytest.mark.parametrize("latex, want_seps", KEEP_WRAPPED)
def test_wrapped_but_result_unchanged(latex, want_seps):
    """会命中 wrap 的三条：空槽仍为 0，但 `sepChr` 会从 `''` 变成真运算符。

    实测（绕过 wrap 直接喂 XSL 对比，`work/p26a/keep.py`）：

        \\left(-x+y\\right)  before 槽 ['−x+y']   sep ['']  →  after 槽 ['−x','y']   sep ['+']
        (-\\frac{a}{b}+1)   before 槽 ['−ab+1']  sep ['']  →  after 槽 ['−ab','1']  sep ['+']
        (-x,\\ y)           before 槽 ['−x,\\xa0y'] sep [''] → after 槽 ['−x','\\xa0y'] sep [',']

    即**修复前是「一个槽装全部、sepChr 空」，修复后是「两个槽、sepChr 是组内真正的
    运算符」** —— 结构确实变了，「sep 不变」是错误的说法。可以接受的理由是渲染等价：
    Word 画「槽1 + sep + 槽2」，与原单槽里的同一串字符观感一致，与 `\\left(a+b\\right)`
    这类天然双槽的写法也一致；不会因为多了个运算符就退化成空方块。

    所以判据是「空槽 0 + sep 等于修复后的真值」，**不是**「与修复前逐字节相同」——
    `want_seps` 填的就是修复后的真值。**wrap 次数同样不是判据。**
    """
    stats: dict = {}
    el = convert(latex, stats)
    assert empty_slots(el) == 0
    assert seps(el) == want_seps, f"{latex!r} 的 sep 与修复后的期望值不符"
    assert stats.get("delim_wraps") == 1


# ---------------------------------------------- 2b. wrap 必须同时包住 body[1]
# 这一组是 M5 变异（只包 body[0]）的唯一抓手。
#
# 空槽数抓不到它：M5 只把「−」包起来，后面那个运算符留在组内当分隔符，于是
# 变成**两个非空槽**（`sep="+"` / `"∓"` / `","` / `"·"`）—— 没有空方块，Word
# 也不报错，只是把 `-(+x)` 画成两个槽。所以要钉的是**槽数与 sep**，不是空槽数。
#
# 为什么这不是「另一种正确排版」：这三种写法里组首的符号都是**一元**的
# （`-+x` 是「负的正的 x」，`-,x` 是「负的 x」），组内并没有真正的分隔语义；
# 掏成两个槽会让 Word 在每个槽里各画一半，与 `(-x)` 的观感也不一致。
WRAP_COVERS_NEXT = [
    r"\left(-+x\right)",
    r"(-+x)",
    r"\left(-\mp x\right)",
    r"(-\mp x)",
    r"\left(-,x\right)",
    r"\left(- \cdot x\right)",
]

@needs_omml
@pytest.mark.parametrize("latex", WRAP_COVERS_NEXT)
def test_wrap_also_swallows_the_next_item(latex):
    """wrap 必须把 body[1] 一起包进来 —— 只包符号会退化成两个槽。

    这一条专门钉变异 M5。改成「只包 body[0]」之后空槽数依然是 0，全套空槽断言
    都绿，只有这里会红：`m:d` 会有 2 个 `m:e` 且 `sepChr` 带回被误判的符号。
    """
    el = convert(latex)
    for delim in el.iter(_m("d")):
        slots = [c for c in delim if c.tag == _m("e")]
        assert len(slots) == 1, (
            f"{latex!r} 的 m:d 有 {len(slots)} 个槽 —— body[1] 没被包进 wrap")
    assert seps(el) == [""], f"{latex!r} 的 sep 不该有值"

# ------------------------------------------------------------------ 3. 13 条边界
BOUNDARY = [
    r"\left(\left(-x\right)\right)",
    r"\left.\frac{a}{b}\right|_{0}^{1}",
    r"\left(\begin{matrix}-a\\b\end{matrix}\right)",
    r"a\left(-b\right)c",
    r"\left(\pm\frac{a}{b}\right)",
    r"\left(\mp x\right)",
    r"\left(-\sqrt{x}\right)",
    r"\left(-\mathbf{v}\right)",
    r"\left(-\text{if }x\right)",
    r"\left(-1\right)^{n}",
    r"\max\{x,y\}\le z",
    r"\frac{dy}{dx}",
    r"\begin{pmatrix}a&b\\c&d\end{pmatrix}",
]

@needs_omml
@pytest.mark.parametrize("latex", BOUNDARY)
def test_boundary_forms_have_no_empty_slot(latex):
    """13 条边界形态修完都是 0 空槽（含矩阵、空定界、取绝对值竖线等）。"""
    el = convert(latex)
    assert empty_slots(el) == 0, f"{latex!r} 仍有空槽"


# ------------------------------------------------------------------ 4. stats
@needs_omml
def test_stats_count_single_double_and_none():
    """`delim_wraps`：单次 = 1、双定界组 = 2、不该动 = 0。"""
    one: dict = {}
    convert(r"(-\frac{y}{x})", one)
    assert one["delim_wraps"] == 1

    two: dict = {}
    convert(r"\left(-\frac{1}{2}\right)\left(+\frac{a}{b}\right)", two)
    assert two["delim_wraps"] == 2

    zero: dict = {}
    convert(r"\left(a+b\right)", zero)
    assert zero.get("delim_wraps", 0) == 0


@needs_omml
def test_stats_accumulate_into_existing_dict():
    """同一份 stats 连着用两次要累加，不是覆盖。"""
    stats: dict = {}
    convert(r"(-x)", stats)
    convert(r"(-y)", stats)
    assert stats["delim_wraps"] == 2


# ------------------------------------------------------------------ 5. 自检
@needs_omml
def test_assert_legal_rejects_empty_slot_in_delim():
    """手搭一个空槽 `m:d`：必须抛 RuntimeError，文案要含症状与两种成因。"""
    from lxml import etree

    el = etree.Element(_m("oMath"))
    delim = etree.SubElement(el, _m("d"))
    dpr = etree.SubElement(delim, _m("dPr"))
    etree.SubElement(dpr, _m("sepChr")).set(_m("val"), "\u2212")
    etree.SubElement(delim, _m("e"))               # 空槽
    body = etree.SubElement(delim, _m("e"))
    etree.SubElement(body, _m("r"))

    with pytest.raises(RuntimeError) as ei:
        math_omml.assert_legal(el, r"\left(-x\right)")
    msg = str(ei.value)
    assert "空槽" in msg, "文案要说清是什么问题"
    assert "sepChr" in msg, "文案要指出是 XSL 的槽分隔符误判"
    assert "left" in msg, "文案里要出现出错公式本身（\\left 由 latex 的 repr 拼入）"


@needs_omml
def test_assert_legal_accepts_fixed_product():
    """真实转换产物（已修）必须过自检 —— 正向用例，防守卫过严。"""
    el = convert(r"(-\frac{y}{x})")
    math_omml.assert_legal(el, r"(-\frac{y}{x})")


@needs_omml
def test_new_check_ignores_rad_and_num_empty_e():
    """`m:rad` 与 `m:num` 下的空 `m:e` 不能被新检查误伤。

    `\\sqrt{}` 本来就产空 `m:rad/m:e`（rad 本身有 e，只是 e 是空的），
    `\\frac{}{}` 产空 `m:num` / `m:den`。这两类都是**别的语义下的空**，
    收成「任何空 m:e 都非法」会把它们打成错误。
    """
    from lxml import etree

    # 1) 真实产物：sqrt{} 与 rac{}{} 都不能抛
    math_omml.assert_legal(convert(r"\sqrt{}"), r"\sqrt{}")
    math_omml.assert_legal(convert(r"\frac{}{}"), r"\frac{}{}")

    # 2) 手搭：m:num 里的空 m:e 也不能抛（与上面真实产物同形，但隔离得更彻底）
    el = etree.Element(_m("oMath"))
    frac = etree.SubElement(el, _m("f"))
    etree.SubElement(frac, _m("num"))
    etree.SubElement(frac, _m("den"))
    math_omml.assert_legal(el, r"\frac{}{}")


@needs_omml
def test_empty_group_still_raises():
    """`\\left(\\right)` 修不好，必须落到自检上抛错 —— 不替它编造内容。

    组内 0 个项、没有 `body[1]` 可包，wrap 无从下手。这一条钉的是「承认修不了」
    而不是「假装修好了」：静默放过它，交付物里就会有一个空方块。
    """
    with pytest.raises(RuntimeError) as ei:
        convert(r"\left(\right)")
    assert "空槽" in str(ei.value)


# ------------------------------------------------------------------ 6. 结构断言
def test_wrap_is_called_by_the_pipeline():
    """`latex_to_omml_lxml` 必须调用 `wrap_leading_unary_in_delim`。

    为什么读源码字符串而不是 mock：这里要钉的是「接线还在」。函数本身有 25 条
    用例覆盖，但**只要有人把那一行调用删掉，那些用例会全部变绿**（没有 wrap 就
    没有空槽检查的输入，`assert_legal` 也不会响）—— 除非正好有别的用例恰好吃到
    空槽。用 AST 反而更脆（改个换行就断），所以直接查 `latex_to_omml_lxml` 的
    源码里有没有这次调用，且**位置必须在 `transform(` 之前** —— 放到之后等于没修，
    MathML 树已经不再被读了。
    """
    import inspect

    src = inspect.getsource(math_omml.latex_to_omml_lxml)
    assert "wrap_leading_unary_in_delim(" in src, (
        "latex_to_omml_lxml 没有再调用 wrap_leading_unary_in_delim —— "
        "定界组空槽修复被摘掉了")
    call_at = src.index("wrap_leading_unary_in_delim(")
    xsl_at = src.index("transform(mml_root)")
    assert call_at < xsl_at, (
        "wrap 必须在 transform 之前跑；放在之后 MathML 树已不再被读，等于没修")


# --------------------------------------------------- 7. docx_validate 空槽扫描
# 这一层刻意**只报不崩**：Word 能打开这些文档（只是画个空方块），所以它不能
# 变成 error，也不能改变退出码 —— `docx_validate` 的 ok/rc 是下游的通过判据，
# 一个「能开但难看」的问题不该把整份文档判死。另外，不含公式的文档**不得出现
# 这个 issue 键**：给所有文档加一条常驻噪音，等于让人学会忽略这份报告。
def _build_docx(path, latex):
    """造一份只有一个段落的 docx；latex 为 None 时不放公式。"""
    from docx import Document
    from docx.oxml import parse_xml

    doc = Document()
    para = doc.add_paragraph("plain text")
    if latex is not None:
        para._p.append(parse_xml(math_omml.latex_to_omml_string(latex)))
    doc.save(str(path))
    return path


@needs_omml
def test_validate_reports_empty_slot_as_warning(tmp_path):
    """含空槽定界组的文档：报 `omml_empty_slot`，但 ok 仍为 True。

    样本必须**绕过**修复链才能造出空槽（`latex_to_omml_lxml` 已经修好了），
    所以这里直接在 document.xml 上动手：拿一条正常公式，再往它的 `m:d` 里塞一个
    空 `m:e` —— 与 XSL 产出的坏形态同构。
    """
    import zipfile

    import docx_validate

    src = _build_docx(tmp_path / "src.docx", r"\left(-x\right)")
    broken = tmp_path / "broken.docx"
    with zipfile.ZipFile(src) as zin:
        items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    with zipfile.ZipFile(broken, "w", zipfile.ZIP_DEFLATED) as zout:
        for item, data in items:
            if item.filename == "word/document.xml":
                xml = data.decode("utf-8")
                marker = "<m:d>"
                assert marker in xml, "样本里没有 m:d，测试前提失效"
                xml = xml.replace(marker, "<m:d><m:e/>", 1)
                data = xml.encode("utf-8")
            zout.writestr(item, data)

    report = docx_validate.validate(str(broken))
    assert report["ok"] is True, "空槽不该让文档判不合法（Word 能打开它）"
    hits = [i for i in report["issues"] if i["code"] == "omml_empty_slot"]
    assert len(hits) == 1, report["issues"]
    assert hits[0]["count"] == 1
    assert hits[0]["severity"] == "warning"


@needs_omml
def test_validate_silent_on_clean_document(tmp_path):
    """修好的文档与纯文字文档都**不得**出现 `omml_empty_slot`。

    两半都要测：只测纯文本文档的话，一个「只要文档里有 m:d 就报」的实现也能过。
    """
    import docx_validate

    fixed = docx_validate.validate(str(_build_docx(tmp_path / "fixed.docx", r"(-\frac{y}{x})")))
    plain = docx_validate.validate(str(_build_docx(tmp_path / "plain.docx", None)))
    for report in (fixed, plain):
        assert report["ok"] is True
        assert not [i for i in report["issues"] if i["code"] == "omml_empty_slot"]
