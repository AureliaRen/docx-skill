# -*- coding: utf-8 -*-
r"""控制序列守卫：LaTeX 专属命令前置硬拒 + OMML 产物残留兜底（P27-A）。

## 这个守卫修的是什么

`latex2mathml` 对**不认识的控制序列不报错**，而是把它当字面量塞进 `<mi>` →
`MML2OMML.XSL` 原样搬进 `<m:t>`。于是交付物的公式里真的印着 \eqno、\textwidth，
而整条链路四个环节全部沉默：转换不抛、`docx_mathify` 的 `stats["failed"]` 只数
异常、退出码只看那个计数、`assert_legal` 只查结构不看文本。取证见
`docs/reports/P27-PROBE-CONTROLSEQ.md`（48 条矩阵：leak 25 / error 1 / dropped 1
/ ok 21）。

## 三层结构，两层判据各自不可省

1. **层 1 前置硬拒**（`assert_no_unsupported_commands`）—— 表里穷举 LaTeX 排版
   专属命令，转换**之前**就抛，能逐条给替代写法。表永远不全（\zzz 列不出来）。
2. **层 2 后置残留**（`assert_no_latex_residue`）—— 扫成品 `m:t`，不看命令名，
   所以任意未知命令都兜得住。**这是 \zzz 的唯一防线**，本文件用「把层 1 打成
   空壳，25 条照样全被层 2 抓到」来钉它，不是靠层 1 捎带。
3. **接入点**（`latex_to_omml_lxml` 里两处调用）—— 用读源码文本的结构断言钉，
   照 `test_math_delim_slots.py::test_wrap_is_called_by_the_pipeline` 的思路：
   函数本身有 25 条用例覆盖，但**只要有人把那一行调用删掉，那些用例会全部变绿**。

## 层 2 的判据为什么是「整段锚定」

残留的**形态**本身就是整段：latex2mathml 把不认识的命令名单独塞进一个 `<mi>`，
XSLT 之后就是一个独立的 `<m:t>`，文本恰好是那条命令。所以判据是
`^\[A-Za-z]+\*?$`。若用「任意位置出现反斜杠」（`\[A-Za-z]+`），`\text{C:\Users}`,
`\text{a\\b}`、`\text{path: \\server\share}` 这三条**完全合法**的公式会被误伤 ——
它们的 `m:t` 里反斜杠是正文内容。三条都钉在下面的 `test_residue_*` 里。

## 上下文判据为什么不能按命令名一刀切

`\nolimits` / `\middle` 在特定位置合法：实测 `\sum\nolimits_{i=1}^{n} a_i` 正常、
`x\nolimits` 泄漏；`\left(a \middle| b\right)` 正常、裸 `a \middle| b` 由
latex2mathml 抛 `ExtraLeftOrMissingRightError`（消息还是空串）。两态各两条钉在
`test_context_*` 里 —— 只测放行那一半，等于没测「扫得准」。

OMML 可用性判定统一走 `tests/_omml.py`（全仓唯一实现）。层 1 在 latex2mathml 之前
就抛，所以那一半不依赖 Office；层 2 要真转出 OMML，`needs_omml` 标记不能省。
"""
from __future__ import annotations

import inspect
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


def _has_latex2mathml() -> bool:
    """层 1 只依赖 latex2mathml（在它之前就抛），不需要 Office。"""
    try:
        import latex2mathml  # noqa: F401
    except Exception:
        return False
    return True


needs_l2m = pytest.mark.skipif(
    not _has_latex2mathml(), reason="needs latex2mathml")


def convert(latex: str, **kw):
    return math_omml.latex_to_omml_lxml(latex, **kw)


def _mtexts(el) -> list[str]:
    return [(t.text or "") for t in el.iter(_m("t"))]


# ===========================================================================
# 判据 1：25 条泄漏逐条抛错，且消息里点名是哪个命令
# ===========================================================================
# 完整输入串照 P27-PROBE-CONTROLSEQ \② 的 leak 25 条，一字不改 —— 自己臆想的
# 输入测不出真问题。tuple 是 (被点名的命令, 输入串)。
LEAKS = [
    (r"\eqno", r"E=mc^2 \eqno{(3-1)}"),
    (r"\label", r"x=1 \label{eq:one}"),
    (r"\nonumber", r"a=b \nonumber"),
    (r"\hfill", r"x=y \hfill"),
    (r"\vspace", r"U=IR \vspace{1em}"),
    (r"\textwidth", r"w=0.5\textwidth"),
    (r"\zzz", r"a \zzz b"),                       # 未知命令：只能靠层 2
    (r"\notag", r"a=b \notag"),
    (r"\eqref", r"x=\eqref{eq:one}"),
    (r"\intertext", r"\intertext{and} x"),
    (r"\shortintertext", r"\shortintertext{and} x"),
    (r"\cline", r"x \cline{1-2}"),
    (r"\multicolumn", r"\multicolumn{2}{c}{x}"),
    (r"\hline", r"a \hline b"),
    (r"\cr", r"a \cr b"),
    (r"\newline", r"a \newline b"),
    (r"\allowbreak", r"a \allowbreak b"),
    (r"\pagebreak", r"a \pagebreak b"),
    (r"\smallskip", r"a \smallskip b"),
    (r"\medskip", r"a \medskip b"),
    (r"\bigskip", r"a \bigskip b"),
    (r"\vspace", r"a \vspace*{1em} b"),            # 带星号的变体同样泄漏
    (r"\qed", r"a \qed"),
    (r"\settowidth", r"\settowidth{\x}{y}"),
    (r"\renewcommand", r"\renewcommand{\x}{1}"),
]


# 24 条由层 1 拒收，只要 latex2mathml 就能跑；\zzz 只能靠层 2，而层 2 要真转出
# OMML（needs_omml = latex2mathml + MML2OMML.XSL）。两条标记分开，否则无 Office 的
# runner 上 \zzz 会以 FileNotFoundError 变红而不是 skip —— CI 实测踩过（PR #8 四腿）。
LEAK_PARAMS = [
    pytest.param(cmd, tex, marks=needs_omml, id=cmd.lstrip(chr(92)) + ":" + tex)
    if cmd == r"\zzz" else
    pytest.param(cmd, tex, id=cmd.lstrip(chr(92)) + ":" + tex)
    for cmd, tex in LEAKS
]


@needs_l2m
@pytest.mark.parametrize("command,latex", LEAK_PARAMS)
def test_every_leak_is_rejected(command, latex):
    r"""每条泄漏都必须抛 RuntimeError，且消息里点名那个命令。

    只断言「抛了」不够：\middle 那条本来就会抛（latex2mathml 的语法异常），
    把「抛 RuntimeError」当唯一判据，会把「守卫没生效、上游恰好报错」也算通过。
    所以必须断言消息里出现**具体命令名** —— 那只有守卫的文案才有。
    """
    with pytest.raises(RuntimeError) as ei:
        convert(latex)
    msg = str(ei.value)
    assert command in msg, (command, msg[:400])


def test_leak_list_is_the_probe_matrix():
    r"""25 条 + 与取证报告 \② 的 leak 列条数一致，且命令名不重复到失真。

    这条防的是「删掉一条难搞的用例让测试变绿」。报告矩阵的 leak 列就是 25 条。
    """
    assert len(LEAKS) == 25
    assert len({c for c, _ in LEAKS}) == 24, "\vspace 故意出现两次（带星号变体）"
    # 参数列表与矩阵同长；\zzz 是唯一带 needs_omml 的那一条。
    assert len(LEAK_PARAMS) == 25
    omml_only = [p for p in LEAK_PARAMS if any(m.name == "skipif"
                                               for m in p.marks)]
    assert len(omml_only) == 1, "只有 \\zzz 需要 XSL，别的条目标错就会在 runner 上少跑"


# ===========================================================================
# 判据 2：21 条 ok 一条都不许误伤（这条比判据 1 更重要）
# ===========================================================================
OK = [
    r"\left( a \middle| b \right)",
    r"\left\{ x \middle| x>0 \right\}",
    r"\frac{a}{b}",
    r"a \hspace{1em} b",                 # 间距被吞（边界），但不该被拒
    r"a \strut b",
    r"\def\foo{1}\foo",                  # 宏定义了**并且**用了：上游会展开
    r"\sum\limits_{i=1}^{n} a_i",
    r"\sum\nolimits_{i=1}^{n} a_i",
    r"\big( x \big)",
    r"\Bigg] x",
    r"\displaystyle \int_0^1 x\,dx",
    r"\textstyle x",
    r"a \phantom{xyz} b",
    r"\vphantom{x} y",
    r"\color{red}{x}",
    r"\textcolor{red}{x}",
    r"x=1 \tag{3-1}",
    r"\overset{a}{b}",
    r"\stackrel{a}{b}",
    r"\mathop{x}",
    r"\substack{a\\b}",
]


@needs_omml
@pytest.mark.parametrize("latex", OK, ids=range(len(OK)))
def test_no_false_positive_on_ok_matrix(latex):
    r"""21 条 ok 全部仍然 ok：不抛、不残留。

    「0 误伤」比「25 条全拒」更重要：把 \sum\nolimits 或 \tag 也毙掉，用户会
    直接弃用这条工具链，而泄漏只是少数公式难看。
    """
    el = convert(latex)
    assert math_omml.find_latex_residue(el) == []


# ===========================================================================
# 判据 3：上下文两态各两条（放行 + 拒，缺一条就等于没测）
# ===========================================================================
@needs_omml
def test_context_nolimits_after_bigop_passes():
    r"""\sum\nolimits_{i=1}^{n} a_i —— 紧跟大算符：放行，且产物干净。"""
    el = convert(r"\sum\nolimits_{i=1}^{n} a_i")
    assert math_omml.find_latex_residue(el) == []


@needs_omml
def test_context_nolimits_without_bigop_is_rejected():
    r"""x\nolimits —— 前面没有大算符：拒，且报的是守卫的文案。

    这是「按命令名一刀切会误伤」的反面证据：同一条 \nolimits，位置一换结论就反。
    """
    with pytest.raises(RuntimeError) as ei:
        convert(r"x\nolimits")
    msg = str(ei.value)
    assert r"\nolimits" in msg
    assert "大算符" in msg, "上下文判据的理由要出现在文案里"


@needs_omml
def test_context_middle_inside_left_right_passes():
    r"""\left(a \middle| b\right) —— 在定界组内：放行。"""
    el = convert(r"\left(a \middle| b\right)")
    assert math_omml.find_latex_residue(el) == []


@needs_omml
def test_context_middle_bare_is_rejected():
    r"""裸 a \middle| b —— 不在 \left…\right 内：拒。

    这条在守卫之前是 latex2mathml 抛 `ExtraLeftOrMissingRightError`，消息是空串，
    用户从 `docx_mathify` 只看到「latex2mathml 转换失败」后面什么都没有。守卫给出
    明确文案（要求：不许变成静默通过）。
    """
    with pytest.raises(RuntimeError) as ei:
        convert(r"a \middle| b")
    msg = str(ei.value)
    assert r"\middle" in msg
    assert "left" in msg, "文案要说清「只能用在 \\left…\\right 里」"


# ===========================================================================
# 判据 4：层 2 独立性 —— \zzz 不在层 1 表里，仍必须被拒
# ===========================================================================
@needs_omml
def test_unknown_command_is_caught_by_layer2_only():
    r"""a \zzz b：层 1 表里查不到 zzz，只能由层 2 抓。

    断言三件事，缺一条就证明不了「层 2 在干活」：
    ① zzz 确实不在层 1 的表里（否则这条是层 1 捎带的）；
    ② 抛的错来自层 2（stats 里 layer == "residue"）；
    ③ 消息点名 \zzz。
    """
    assert "zzz" not in math_omml.LATEX_ONLY_COMMANDS
    assert "zzz" not in math_omml.CONTEXT_RULES
    assert math_omml.find_unsupported_commands(r"a \zzz b") == []

    stats: dict = {}
    with pytest.raises(RuntimeError) as ei:
        convert(r"a \zzz b", stats=stats)
    assert r"\zzz" in str(ei.value)
    assert stats["latex_guard"] == 1
    assert stats["latex_guard_hits"][0]["layer"] == "residue"


@needs_omml
def test_layer2_covers_all_25_leaks_without_layer1(monkeypatch):
    r"""把层 1 打成空壳后，25 条泄漏**照样**全被层 2 抓到。

    这是「层 2 不是层 1 捎带」的硬证据：关掉层 1 再走一遍完整转换链，只看
    `find_latex_residue` 的读数。任何一条在新判据下漏掉，都说明它的残留形态不是
    「整段一条命令」—— 那是判据问题，必须报出来，不许改判据去凑。
    """
    monkeypatch.setattr(math_omml, "assert_no_unsupported_commands",
                        lambda latex, stats=None: None)
    monkeypatch.setattr(math_omml, "assert_no_latex_residue",
                        lambda omml, latex, stats=None: None)
    missed = []
    for command, latex in LEAKS:
        el = math_omml.latex_to_omml_lxml(latex)
        if not math_omml.find_latex_residue(el):
            missed.append((command, latex, _mtexts(el)))
    assert missed == [], missed


# ===========================================================================
# 判据 4b：编号 / 引用命令的**作用域**（整篇链路 vs 单条入口，两态都要）
# ===========================================================================
# 这三条命令在本表里对**不同入口**地位完全不同，必须用两态测出来，不能只测一半：
#
# * **整篇链路**：docx_mathify 的 `_mark_commands()`（:498）在 `_analyze` **之前**
#   就把 \label{} / \eqref{} / \ref{} 摘成 special atom 做编号，所以它们
#   **根本到不了** latex_to_omml_lxml —— 层 1 在这里是死代码（拦不到、也不必拦）。
# * **单条入口**：insert_formula / docx_create 的 latex block 没有那层预处理，命令
#   一路走到 latex2mathml 变字面量 —— 层 1 是**唯一防线**。
#
# 只测后一半会让人误以为「整篇链路也靠它拦」；只测前一半会以为「这条守卫没用」。
REF_CMDS = [(r"\label", r"x=1 \label{eq:one}"),
            (r"\eqref", r"x=\eqref{eq:one}"),
            (r"\ref", r"a \ref{eq:one}")]


@pytest.mark.parametrize("command,latex", REF_CMDS, ids=[c.lstrip(chr(92)) for c, _ in REF_CMDS])
def test_mark_commands_strips_ref_commands_before_conversion(command, latex):
    r"""整篇链路：`_mark_commands` 把这些命令摘成编号原子，公式片段里不再有它们。

    这条钉的是「层 1 在整篇链路上是死代码」这个**事实** —— 不是抱怨。哪天有人把
    `_mark_commands` 拆了或改了顺序，这里会红，因为那条命令会重新活到转换器面前。
    """
    import docx_mathify

    atoms = [{"ch": ch, "rpr": None, "script": None} for ch in latex]
    marked = docx_mathify._mark_commands(atoms)
    joined = "".join(a["ch"] for a in marked)
    assert command not in joined, (
        "%s 仍留在公式片段里 —— _mark_commands 没有摘掉它，整篇链路会把它送进转换器"
        % command)
    specials = [a["special"] for a in marked if a.get("special")]
    assert specials, "没有被摘成 special atom"
    assert specials[0][0] == command.lstrip(chr(92))


@needs_l2m
@pytest.mark.parametrize("command,latex", REF_CMDS, ids=[c.lstrip(chr(92)) for c, _ in REF_CMDS])
def test_single_entry_rejects_ref_commands(command, latex):
    r"""单条入口：`insert_formula.build_para` 没有预处理，层 1 是唯一防线。

    走**真实**的 build_para（不是直接调 assert），因为要证明的是「这条入口上确实
    有人拦」；直接调 assert 会把「build_para 有没有走 latex_to_omml_lxml」这个
    前提省掉，而那个前提正是覆盖率论证的核心。
    """
    import insert_formula

    with pytest.raises(RuntimeError) as ei:
        insert_formula.build_para(latex, False)
    msg = str(ei.value)
    assert command in msg, (command, msg[:300])
    assert "单条入口" in msg, "文案要说清这一层拦的是哪条入口"


@needs_l2m
def test_ref_commands_pass_through_mathify_end_to_end():
    r"""整篇链路端到端：带 \label 的源文本**不会**被守卫拒收（_mark_commands 先摘了）。

    与前两条互补：那条说「_mark_commands 摘掉了」，这条说「所以整篇链路照常出公式」。
    如果哪天守卫开始在整篇链路上拒绝 \label，这条会红 —— 而那是个回归：编号是
    docx_mathify 的功能，不该被公式守卫毙掉。
    """
    import docx_mathify

    atoms = [{"ch": ch, "rpr": None, "script": None}
             for ch in r"x=1 \label{eq:one}"]
    _segs, _maths, _eqnum, labels, _align = docx_mathify._analyze(
        docx_mathify._mark_commands(atoms), False)
    assert "eq:one" in (labels or []), "标签没被收集，编号链路的入口断了"


# ===========================================================================
# 判据 4c：\limits 的判据**必要性**（摘掉它就静默泄漏）
# ===========================================================================
@needs_omml
def test_limits_rule_is_necessary(monkeypatch):
    r"""把 `limits` 从 CONTEXT_RULES 摘掉，`\text{a\limits}` 就会**静默泄漏**。

    ## 为什么这条必须有（而不是「留着只为口径一致」）

    实测两态：

    * 顶层 `x\limits` 被 latex2mathml **静默吃掉**（m:t 只有 x）—— 不泄漏；
    * 但 `\text{a\limits}` 会跟着正文一起变字面量，m:t 是 a 加反斜杠加 limits。

    而层 2 是整段锚定（整条 m:t 必须就是一条命令），整段是 a 加反斜杠加 limits、
    不是单条命令，所以**看不见**这一条。也就是说：有层 2 的这个盲区，`limits` 这条
    判据就不是冗余 —— 摘掉它，那条公式会静默带着反斜杠落进文档。

    这条用例本身就是在做那个实验：临时删除条目 → 断言行为退化成静默泄漏 → 复原。
    """
    latex = r"\text{a\limits}"

    # 现状：层 1 拒收
    with pytest.raises(RuntimeError) as ei:
        convert(latex)
    assert r"\limits" in str(ei.value)

    # 实验：摘掉条目后，光靠层 2 抓不到（→ 静默泄漏）
    saved = math_omml.CONTEXT_RULES.pop("limits")
    try:
        el = math_omml.latex_to_omml_lxml(latex)
        assert _mtexts(el) == [r"a\limits"], _mtexts(el)
        assert math_omml.find_latex_residue(el) == [], (
            "层 2 竟然抓到了 —— 那 limits 判据确实冗余，应该删掉；"
            "现在的实测是它抓不到，所以条判据必要")
    finally:
        math_omml.CONTEXT_RULES["limits"] = saved
    assert "limits" in math_omml.CONTEXT_RULES, "复原失败"


@needs_omml
def test_limits_still_allowed_after_bigop():
    r"""`\sum\limits_{i=1}^{n} a_i` 放行（上下文判据的另一半，别只测拒收）。"""
    el = convert(r"\sum\limits_{i=1}^{n} a_i")
    assert math_omml.find_latex_residue(el) == []


@needs_l2m
def test_limits_bare_is_rejected():
    r"""`x\limits`（不在大算符后）拒收 —— 与上一条对照，同一个命令两种结论。"""
    with pytest.raises(RuntimeError) as ei:
        convert(r"x\limits")
    msg = str(ei.value)
    assert r"\limits" in msg
    assert "大算符" in msg

# ===========================================================================
# 判据 5：层 1 独立性 —— 压根没进 latex2mathml
# ===========================================================================
@needs_l2m
def test_layer1_stops_before_latex2mathml(monkeypatch):
    r"""层 1 命中时 `latex2mathml.converter.convert` 的调用次数必须是 **0**。

    证据价值：只断言「抛了 RuntimeError」分不清是守卫拦下的，还是 latex2mathml
    自己语法错、或者 XSL 失败被包成了同一个异常类型。计数为 0 才证明「压根没进
    转换器」，也就是层 1 真在转换**之前**。
    """
    import latex2mathml.converter as lconv

    calls = []
    real = lconv.convert

    def counting(*a, **kw):
        calls.append(a)
        return real(*a, **kw)

    monkeypatch.setattr(lconv, "convert", counting)
    with pytest.raises(RuntimeError):
        convert(r"E=mc^2 \eqno{(3-1)}")
    assert calls == [], "层 1 命中后不该再去调 latex2mathml"

    # 对照：合法公式必须真走转换器，否则上一条的「0 次」可能是空转出来的
    math_omml.assert_no_unsupported_commands(r"\frac{a}{b}")
    assert calls == []


# ===========================================================================
# 判据 6：接入点结构断言（防有人把调用删了、测试还绿）
# ===========================================================================
def test_both_guards_are_called_by_the_pipeline():
    r"""`latex_to_omml_lxml` 必须**同时**调用两层守卫，且位置正确。

    位置不是洁癖：层 1 若在 `keep_spacing()` 之后，间距命令已被改写成控制空格，
    扫出来的位置和命令都不对；层 2 若在 `split_runs_for_upright()` 之前，扫的是
    中间态，拆 run 之后还会再产出一批 m:t —— 那时漏掉的就查不回来了。
    """
    src = inspect.getsource(math_omml.latex_to_omml_lxml)
    assert "assert_no_unsupported_commands(" in src, (
        "latex_to_omml_lxml 没有再调用 assert_no_unsupported_commands —— 前置硬拒被摘掉了")
    assert "assert_no_latex_residue(" in src, (
        "latex_to_omml_lxml 没有再调用 assert_no_latex_residue —— 残留兜底被摘掉了")

    pre_at = src.index("assert_no_unsupported_commands(")
    spacing_at = src.index("keep_spacing(latex)")
    assert pre_at < spacing_at, (
        "层 1 必须在 keep_spacing 之前扫原始参数；之后扫的是被改写过的串")

    split_at = src.index("split_runs_for_upright(")
    residue_at = src.index("assert_no_latex_residue(")
    assert split_at < residue_at, (
        "层 2 必须在 split_runs_for_upright 之后扫最终树；之前扫的是中间态")


def test_guard_tables_are_consistent():
    r"""表本身的自洽性：上下文表里的 rule 都有实现，命令名都不带反斜杠。

    这条防的是「加了 CONTEXT_RULES 但忘了往 _CONTEXT_PREDICATES 里挂实现」——
    那种情况下 `_context_allows` 会 KeyError，而只有刚好测到那条命令的用例才会红。
    """
    for name in math_omml.LATEX_ONLY_COMMANDS:
        assert not name.startswith(chr(92)), name
        assert name.isalnum() or name.isalpha(), name
        assert math_omml.LATEX_ONLY_COMMANDS[name].strip(), name
    for name, info in math_omml.CONTEXT_RULES.items():
        assert info["rule"] in math_omml._CONTEXT_PREDICATES, name
        assert info["reason"].strip() and info["suggestion"].strip(), name
    # 两条表不许重叠：同一条命令既「绝对不行」又「看情况」会把判据搅成玄学
    assert not (set(math_omml.LATEX_ONLY_COMMANDS)
                & set(math_omml.CONTEXT_RULES))


# ===========================================================================
# 判据 7：误报语料 —— 复杂但完全合法的公式必须一条不拦
# ===========================================================================
@needs_omml
def test_residue_ignores_backslashes_inside_text():
    r"""`\text{...}` 里的反斜杠是**正文**，不是残留。

    这三条是层 2 判据的成败线：用「任意位置出现反斜杠」的判据会把它们全误伤。
    `\text{C:\Users\me}` 在 m:t 里是 `C:\Users\me`，一个 Windows 路径；
    `\text{a\\b}` 是用户想要的**一个反斜杠字形**；`\text{path: \\server\share}`
    是 UNC 路径。三条都是完全合法、必须放行的公式。
    """
    for latex in (r"\text{C:\Users\me}",
                  r"\text{a\\b}",
                  r"\text{path: \\server\share}"):
        el = convert(latex)
        assert math_omml.find_latex_residue(el) == [], (latex, _mtexts(el))


@needs_omml
def test_residue_catches_control_sequence_inside_text():
    r"""文本模式：`\text{a \zzz b}` **必须**被层 2 抓到（整段锚定抓不到）。

    这条是「按空白切词」判据存在的全部理由。实测形态：latex2mathml 在 `\text{}` 里
    用 NBSP（U+00A0）填空白，于是 m:t 是 `a` + NBSP + `\zzz` + NBSP + `b` ——
    整段不是单条命令，旧的整段锚定判据看不见它，命令就静默落进正文。

    同时断言「这不是层 1 捎带的」：`zzz` 不在层 1 的任何表里，所以拦截只能来自层 2，
    且 stats 的 layer 必须是 `"residue"`。
    """
    latex = r"\text{a \zzz b}"
    assert math_omml.find_unsupported_commands(latex) == [], "层 1 不该命中 zzz"

    stats: dict = {}
    with pytest.raises(RuntimeError) as ei:
        convert(latex, stats=stats)
    assert r"\zzz" in str(ei.value)
    assert stats["latex_guard_hits"][0]["layer"] == "residue"


@needs_omml
def test_residue_known_blind_spot_glued():
    r"""已知边界：`\text{a\zzz b}`（反斜杠**紧贴**字母）仍漏 —— 钉住，别当 bug。

    m:t 是 `a\zzz` + NBSP + `b`，切出来的 token 是 `a\zzz` —— 不是单条命令。
    抓它必须回到「任意位置匹配」，而那样 `\text{C:\Users\me}` 这类路径公式会
    全被误伤（三条护栏见 `test_residue_ignores_backslashes_inside_text`）。这个取舍是
    刻意的：宁可漏「紧贴的未知命令」，也不误杀正常正文。

    这条测试的用途是**防误解**：下一个人看到漏，先来这里确认这是记录在案的边界。
    """
    latex = r"\text{a\zzz b}"
    el = convert(latex)
    assert _mtexts(el) == ["a" + r"\zzz" + "\u00a0b"], _mtexts(el)
    assert math_omml.find_latex_residue(el) == [], "紧贴形态命中就意味着判据改成任意位置了"


@needs_omml
def test_residue_path_guard_survives_token_split():
    r"""三条路径护栏在「按空白切词」下仍然 0 误报（换了判据之后必须重验）。

    `\text{\Users\me}` 是「反斜杠开头」的正文，最容易在切词后误伤 ——
    它的 token 是 `\Users\me`，因为 `\Users` 后面紧跟 `\me`（无反斜杠前空白），
    整条不是 `^\[A-Za-z]+\*?$`。三条一起钉，少一条就少一个方向的保证。
    """
    for latex in (r"\text{C:\Users\me}",
                  r"\text{\Users\me}",
                  r"\text{a\\b}"):
        el = convert(latex)
        assert math_omml.find_latex_residue(el) == [], (latex, _mtexts(el))


@needs_omml
def test_residue_requires_command_token():
    r"""判据是「**以空白为界的词**本身是一条命令」，不是「整段」、也不是「含反斜杠」。

    用文本手搓的 `m:t` 造三个方向（不依赖转换链，所以不受上游改行为影响）：

    * `\zzz` 单独一段 → 命中；
    * `C:\zzz` 带前缀 → 不命中（那是正文内容）；
    * `a` + NBSP + `\zzz` + NBSP + `b`（文本模式的真实形态）→ **命中** ——
      这正是「整段锚定」抓不到、而「按空白切词」能抓到的那个洞。
    """
    from lxml import etree

    def omml_with(*texts):
        root = etree.Element(_m("oMath"), nsmap={"m": M})
        for t in texts:
            run = etree.SubElement(root, _m("r"))
            etree.SubElement(run, _m("t")).text = t
        return root

    assert math_omml.find_latex_residue(omml_with(r"\zzz")) == [r"\zzz"]
    assert math_omml.find_latex_residue(omml_with("C:" + r"\zzz")) == []
    assert math_omml.find_latex_residue(omml_with("ab")) == []
    # 文本模式的真实形态：命令夹在正文之间，用 NBSP（U+00A0）分隔 —— 命中
    assert math_omml.find_latex_residue(
        omml_with("a\u00a0" + r"\zzz" + "\u00a0b")) == [r"\zzz"]
    # 紧贴字母（无反斜杠前空白）→ token 是 a\zzz，不是单条命令 → 不命中（已知边界）
    assert math_omml.find_latex_residue(omml_with("a" + r"\zzz" + "\u00a0b")) == []
    # 只认 m:t：w:t 里的反斜杠是合法正文（正文写路径），不能报
    w_root = etree.Element("{%s}p" % math_omml.W_NS,
                           nsmap={"w": math_omml.W_NS})
    etree.SubElement(w_root, "{%s}t" % math_omml.W_NS).text = r"\Users\me"
    assert math_omml.find_latex_residue(w_root) == []


COMPLEX_OK = [
    r"f(x) = \begin{cases} x^{2}, & x \ge 0 \\ -x, & x < 0 \end{cases}",
    r"A = \begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix}",
    r"\begin{pmatrix} \frac{a}{b} & 0 \\ 0 & 1 \end{pmatrix}",
    r"\int_0^1 f(x)\,\mathrm{d}x",
    r"\lim_{x\to 0} \frac{\sin x}{x} = 1",
    r"\dot{I}_{L} + \dot{I}_{C}",
    r"\hat{U} = \vec{F} \cdot \overline{z}",
    r"f(x) = \begin{cases} 1, & \text{if }x>0 \\ 0, & \text{otherwise} \end{cases}",
    r"\boldsymbol{\alpha} + \partial u / \partial x",
    r"\nabla \cdot \vec{E} = \frac{\rho}{\varepsilon_0}",
    r"x \to \infty",
    r"a \pm b \mp c",
    r"\sqrt[3]{x} + \binom{n}{k}",
    r"\left\| x \right\| + \left\{ y \right\}",
    r"\sum_{i=1}^{n} a_i = \prod_{j=1}^{m} b_j",
    r"\frac{\partial^2 u}{\partial x^2}",
    r"\begin{aligned} f &= x^{2} \\ &= y \end{aligned}",
    r"\mathrm{tr}(A) = \lambda_{1} + \lambda_{2}",
]


@needs_omml
@pytest.mark.parametrize("latex", COMPLEX_OK, ids=range(len(COMPLEX_OK)))
def test_safe_entry_accepts_legal_complex_formulas(latex):
    r"""走 `latex_to_omml_safe`（调用方的默认入口）逐条不抛。

    18 条合法复杂公式：环境（cases / matrix / aligned）、分式嵌套、极限、相量点、
    重音、希腊字母、偏导、∇、∞、±∓、根式、二项式、定界组、求和连乘。
    这条防「守卫太凶把好公式也毙了」—— 那比漏掉 \eqno 严重得多。
    """
    el = math_omml.latex_to_omml_safe(latex)
    assert el is not None
    assert math_omml.find_latex_residue(el) == []


# ===========================================================================
# 判据 8：stats 透传口径（拒收时留什么、不留什么，写清并测死）
# ===========================================================================
@needs_l2m
def test_stats_on_rejection_records_guard_not_failed():
    r"""拒收记 `latex_guard`，**不**记 `failed` —— 两者是不同的问题。

    `failed` 的语义是「公式没转成」（docx_mathify 用它算退出码 3）。拒收是「源文本
    里有 GOML 表达不了的命令」，用户要改的是源文本，不是重试转换器。混进 failed
    会让「一条命令写错」看起来像「工具坏了」。
    """
    stats: dict = {"failed": []}
    with pytest.raises(RuntimeError):
        convert(r"a=b \nonumber", stats=stats)
    assert stats["latex_guard"] == 1
    hit = stats["latex_guard_hits"][0]
    assert hit["layer"] == "pre"
    assert hit["commands"] == [r"\nonumber"]
    assert stats["failed"] == [], "拒收不许污染 failed（那是退出码 3 的依据）"


@needs_omml
def test_stats_on_success_has_no_guard_keys():
    r"""成功时不留 `latex_guard` / `latex_guard_hits`（不给报告挂常驻噪音）。

    与 `docx_validate` 的 `omml_empty_slot` 同一条理由：键常驻就等于教人忽略它。
    """
    stats: dict = {}
    convert(r"\frac{a}{b}", stats=stats)
    assert "latex_guard" not in stats
    assert "latex_guard_hits" not in stats


@needs_l2m
def test_stats_none_does_not_crash():
    r"""不传 stats 时三层行为完全一样，只是不留痕（与既有 stats 口径一致）。"""
    with pytest.raises(RuntimeError):
        convert(r"a=b \notag")


# ===========================================================================
# 判据 9：已知边界 —— 本轮不修，但现状必须被钉住（防「以为修了」）
# ===========================================================================
@needs_omml
def test_boundary_hspace_still_swallowed():
    r"""`\hspace{1em}` 的间距仍被吞（不泄漏、也不报错）—— 本轮不修。

    它与泄漏是**两种**失真：泄漏有反斜杠可扫，吞间距只能比对源 LaTeX 与 m:t 的
    文本量。守卫不接这一类，钉在这里是为了让下一个人知道「这不是修好了」。
    """
    el = convert(r"a \hspace{1em} b")
    assert _mtexts(el) == ["a", "b"], "间距仍然被吞：本轮不修"
    assert math_omml.find_latex_residue(el) == []


@needs_omml
def test_boundary_rule_still_produces_empty_formula():
    r"""`\rule{1em}{2pt}` 仍退化成空公式（dropped）—— 本轮不修。

    没人能扫出「本该有内容却空着」：空 m:oMath 是合法 XML，Word 打开也不报错。
    要抓它得比对源 LaTeX 与产物文本量，属另一类问题。
    """
    el = convert(r"\rule{1em}{2pt}")
    assert [t for t in _mtexts(el) if t.strip()] == [], "仍然是一条空公式：本轮不修"
    assert math_omml.find_latex_residue(el) == []
