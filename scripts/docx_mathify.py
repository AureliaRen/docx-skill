# -*- coding: utf-8 -*-
r"""把整篇 docx 里的「线性公式」批量换成 LaTeX 渲染的原生 OMML 公式。

适用对象：公式写成「Unicode 字符 + 真上下标 run」的文档（`U`+下标`oc`、`P∝I`+上标`2`、
`√2`、`∠53.1°`、`1kΩ=10³Ω`）。中文正文、表格结构、字体字号全部原样保留。

转换内容：
    U+下标 oc        → U_{oc}                上下标 run → ^{} / _{}
    a/b  (6/(2+2))   → \frac{a}{b}           仅两侧都是完整操作数时转分式
    √2 / √(LC)       → \sqrt{2} / \sqrt{LC}
    ω φ ψ τ Ω Δ Σ ∠ ∝ ≈ ≤ × · °  → \omega … \angle \propto \approx \leq \times \cdot ^{\circ}
    U + 组合点       → \dot{U}
    sin cos tan ln   → \sin \cos \tan \ln
    末尾单位 1.5A    → 1.5\,\text{A}
    单位斜杠 kg/m³   → \mathrm{kg/m^{3}}       （`/` 读作「每」，不转分式）
    真分式   a/b     → \frac{a}{b}            （两侧是量符号 / 数字，`/` 才是分数线）

有意不转（避免造出半个公式）：纯数量（`6V`、`2Ω`）、残缺片段（`b=`、`=R`、`Uoc →`）、日期。

量纲串判定见 `linear_to_latex._dimensionify()`：`/` 两侧都是单位符号只是必要条件，
还要有「数字锚点」（`1.2kg/m³`）或命中显式登记的复合单位（`mol/L`）。
`v=s/t`、`p=F/A`、`ρ=m/V` 这类真分式因此不会被误读成单位。
报告里会列出本段识别到的量纲串，便于人工核对。

附带修正：源文档把「后面的正常字符」也设成上下标时（`I^{2R}` / `R_{1R2}` / `ω_{0L}`），
自动拆成 `I^{2}R` / `R_{1}R_{2}` / `ω_{0}L`，并在报告里列出改了什么。

Usage
-----
    python docx_mathify.py in.docx -o out.docx
    python docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex
    python docx_mathify.py in.docx --inplace --no-merged-script-fix
    python docx_mathify.py in.docx -o out.docx --dry-run     # 只出报告，不写 docx

打开输出后建议在真 Word 里验一遍：
    python verify_word.py out.docx --out verify        # latex-fix 工作区的校验脚本
"""
from __future__ import annotations

# --- repo bootstrap: make sibling imports work from any cwd ---------------
import _bootstrap  # noqa: F401 — UTF-8 stdio on Windows
# --------------------------------------------------------------------------

import argparse
import re
import shutil
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Twips
from docx.text.paragraph import Paragraph

import linear_to_latex as L
from math_omml import latex_to_omml_lxml
from fix_phasor_dot import fix_phasor_dots

XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

# --------------------------------------------------------------- 混合角标策略
# 主人 2026-11-06 的判据（HANDOFF-NEXT §一.2）：
#   「选中角标，按 Ctrl+Shift+= 后**它应该返回到正常的大小**，这才是真的角标」
#
# OMML 的 `m:sSup` 做不到这一点 —— 它是公式引擎排出来的小号字，
# Word 的 `Font.Superscript` 恒为 0，再按一次上标只会在小号字上再叠一层。
# 真角标是 run 级的 `w:rPr/w:vertAlign`，Word 里 `Font.Superscript == -1`
# 且字号不变、`OMaths.Count == 0`。
#
# 所以按「混合策略」分流（路线 A）：
#   * **纯上下标 span**（只有字母/数字/希腊字母/运算符，没有分式、根号、积分、
#     求和、矩阵、多行环境、绝对值）→ 不产 OMML，直接落成普通 run + vertAlign。
#   * **真结构公式**（`\frac`、`\sqrt`、`\int`、`\sum`、`\begin{…}`、`\left|`）
#     → 仍走 `latex_to_omml_lxml`，行为一点不变。
#
# 模块级开关，置 False 即完全回到「全部 OMML」的旧行为（便于回退/对拍）。
MIXED_SUPERSCRIPT = True

# --------------------------------------------------------------- 免排版分流
# 主人 2026-11-07：「我发现很简单的完全不需要 omml 的，比如 `s=10km` 这种，
# 也会 LaTeX 转成 OMML。」
#
# 上一轮的混合策略只救了「带角标的」—— `is_pure_script` 要求片段里**至少有一个**
# `^`/`_`。所以连角标都没有的简单等式（`s=10km` / `t=2h` / `m=5kg`）一路落到
# OMML 分支，可在 Word 里点进去是公式编辑器，量符号被排成数学斜体、单位 `km`
# 也跟着斜，改一个字要开公式工具。
#
# 这类片段**没有任何二维排版需求**（无分式/根号/上下标），落成普通文字 run 才是
# 它该有的样子。判据见 `no_layout_reason()`。量符号因此不再保数学斜体 ——
# 主人 2026-11-07 已拍板接受这个代价（PLAN-ROUND4 §0）。
#
# 模块级开关，置 False 即完全回到「全部 OMML」的旧行为（便于回退/对拍）。
# 判据函数本身保持纯函数，开关只在调用点（`_om_or_scripts`）生效。
NO_LAYOUT_NEEDED = True

# 出现这些就**不能**降级成文字角标：它们有二维排版（上下分数、根号覆盖、
# 上下限、矩阵），文字 run 表达不了。`\mathrm{}` / `\text{}` 不在此列 ——
# 它们只是正体字，一维的，可以安全展开成普通字符。
_STRUCTURAL_CMDS = (
    r"\frac", r"\dfrac", r"\tfrac", r"\cfrac", r"\binom", r"\genfrac",
    r"\sqrt", r"\surd",
    r"\int", r"\oint", r"\iint", r"\iiint",
    r"\sum", r"\prod", r"\coprod",
    r"\lim", r"\limsup", r"\liminf",
    r"\overline", r"\underline", r"\overbrace", r"\underbrace",
    r"\vec", r"\hat", r"\bar", r"\dot", r"\ddot", r"\tilde",
    r"\stackrel", r"\overset", r"\underset",
    r"\left", r"\right", r"\middle",
    r"\begin", r"\end", r"\matrix", r"\array", r"\substack",
    r"\tag", r"\label", r"\ref", r"\eqref",
)

# Unicode → LaTeX 的反向表（linear_to_latex 的 GREEK/SYMBOL 是正向的）。
# 只收「一维、可直接落成字符」的映射；`\frac{1}{2}` 这类结构式输出到不了这里
# （`\frac` 已在 _STRUCTURAL_CMDS 里被拦掉）。
# 显式间距落成的字符：NBSP。原因见 _PLAIN_CMD 里 `\,` / `\;` 那段注释。
_NBSP = "\u00a0"

_PLAIN_CMD: dict[str, str] = {cmd: ch for ch, cmd in L.GREEK.items()}
for _ch, _cmd in L.SYMBOL.items():
    if not _cmd.startswith("\\") or "{" in _cmd or "}" in _cmd:
        continue
    _PLAIN_CMD.setdefault(_cmd, _ch)
_PLAIN_CMD.update({
    # —— 显式间距命令：作者有意留的空，必须落成真空格 ——
    # `\,` / `\:` / `\;` 是细空、中空、粗空；落成 NBSP（U+00A0）而不是普通空格：
    # ① 所有字体都有这个字形（U+2009 THIN SPACE 在 SimHei 里可能缺，静默变方框）；
    # ② 不断行，`18 g/mol` 的数值与单位不会被拆到两行。
    # 之前这里映射成 ""，等于把作者的显式间距**吃掉**了（`M=18\,\text{g/mol}`
    # 会渲染成 `M=18g/mol`）。
    r"\,": _NBSP, r"\:": _NBSP, r"\;": _NBSP,
    r"\quad": _NBSP, r"\qquad": _NBSP + _NBSP,
    r"\!": "",                      # 负细空：本来就是「往回缩」，落成空串
    r"\ ": " ", r"~": _NBSP,        # `~` 在 LaTeX 里就是不断行空格
    r"\%": "%", r"\&": "&", r"\#": "#",
    r"\_": "_", r"\^": "^", r"\$": "$",
    r"\{": "{", r"\}": "}",
    r"\lvert": "|", r"\rvert": "|", r"\vert": "|",
    r"\langle": "\u27e8", r"\rangle": "\u27e9",
    # `30^{\circ}` 里的 `\circ`：度符号是单字符，落成 ° 就够（它在上标里）
    r"\circ": "\u00b0",
})
# `\mathrm{...}` / `\text{...}`：正体组，内容照常展开（不保留命令）
_PLAIN_GROUP_CMDS = (r"\mathrm", r"\text", r"\textrm", r"\operatorname")

# 函数名是一维正体文字（`\sin30°` 里只有 ° 需要上标），可以安全降级；
# 它们**不是**结构（没有上下限），所以不进 _STRUCTURAL_CMDS。
# 但 `\lim`/`\max`/`\sum` 这类带上下限的留在结构表里。
for _fn in ("arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
            "sin", "cos", "tan", "cot", "sec", "csc",
            "ln", "lg", "log", "exp", "arg"):
    _PLAIN_CMD.setdefault("\\" + _fn, _fn)


def _group_end(s: str, i: int) -> int:
    """s[i] == '{'，返回配对 '}' 之后的下标；不配对返回 -1。"""
    depth = 0
    for j in range(i, len(s)):
        if s[j] == "{":
            depth += 1
        elif s[j] == "}":
            depth -= 1
            if depth == 0:
                return j + 1
    return -1


def structural_reason(latex: str) -> str | None:
    """片段里第一个「二维排版结构」的命令名；没有则 None。

    `\\int` 只在**命令边界**成立：`\\infty` 里也有 `\\in` 前缀，但不该被误判。
    """
    for cmd in _STRUCTURAL_CMDS:
        start = 0
        while True:
            i = latex.find(cmd, start)
            if i < 0:
                break
            tail = latex[i + len(cmd):i + len(cmd) + 1]
            if not tail.isascii() or not tail.isalpha():
                return cmd                     # 命令边界 → 真是这个结构
            start = i + 1                      # 只是更长命令名的前缀，跳过
    return None


def is_pure_script(latex: str) -> bool:
    r"""这个 LaTeX 片段能不能降级成「普通文字 + 真角标」。

    两条同时成立才算：

    1. **没有任何二维排版结构**（`_STRUCTURAL_CMDS` 全不出现）；
    2. **至少带一个** `^` 或 `_` —— 没有角标的片段（`y=f(x)`、`A=(-1,3)`）
       转了也白转，只会把 Cambria Math 换成正文字体，制造无谓的差异。

    判据刻意偏保守：宁可留 OMML，也不要把分式降级成一串字符。
    """
    return is_pure_script_detail(latex)[0]


def is_pure_script_detail(latex: str) -> tuple[bool, str]:
    """同 `is_pure_script`，但连**理由**一起返回（探针与报告要可核对）。"""
    if not latex:
        return False, "empty"
    if "^" not in latex and "_" not in latex:
        return False, "no_script"
    reason = structural_reason(latex)
    if reason is not None:
        return False, "structural:" + reason
    if _latex_to_pieces(latex) is None:
        return False, "unexpandable_cmd"
    return True, "pure_script"


def _latex_to_pieces(latex: str) -> list[tuple[str, str | None]] | None:
    r"""LaTeX → ``[(文字, vertAlign 或 None)]``；遇到不认识的命令返回 None。

    *vert* 取 ``"superscript"`` / ``"subscript"`` / ``None``（基线）。
    只处理一维内容 —— 调用方已用 `is_pure_script` 挡掉结构式。
    """
    out: list[tuple[str, str | None]] = []

    def emit(text: str, vert: str | None) -> None:
        if not text:
            return
        if out and out[-1][1] == vert:
            out[-1] = (out[-1][0] + text, vert)
        else:
            out.append((text, vert))

    def walk(s: str, vert: str | None) -> bool:
        i, n = 0, len(s)
        while i < n:
            ch = s[i]
            # —— 角标：^ / _ 后接 {组} 或单字符 ——
            if ch in "^_":
                want = "superscript" if ch == "^" else "subscript"
                i += 1
                if i >= n:
                    return False
                if s[i] == "{":
                    end = _group_end(s, i)
                    if end < 0:
                        return False
                    if not walk(s[i + 1:end - 1], want):
                        return False
                    i = end
                else:
                    emit(s[i], want)
                    i += 1
                continue
            # —— 分组：只是作用域，内容照常展开 ——
            if ch == "{":
                end = _group_end(s, i)
                if end < 0:
                    return False
                if not walk(s[i + 1:end - 1], vert):
                    return False
                i = end
                continue
            # —— 命令 ——
            if ch == "\\":
                m = re.match(r"\\([A-Za-z]+|.)", s[i:])
                if not m:
                    return False
                name = "\\" + m.group(1)
                j = i + m.end()
                if name in _PLAIN_GROUP_CMDS:
                    while j < n and s[j] == " ":
                        j += 1
                    if j < n and s[j] == "{":
                        end = _group_end(s, j)
                        if end < 0:
                            return False
                        if not walk(s[j + 1:end - 1], vert):
                            return False
                        j = end
                    elif j < n:                 # `\mathrm m` 这种无组写法
                        if not walk(s[j], vert):
                            return False
                        j += 1
                elif name in _PLAIN_CMD:
                    emit(_PLAIN_CMD[name], vert)
                else:
                    return None                 # 未知命令：交给 OMML
                # ★ 控制字（纯字母命令，如 `\times`、`\infty`、`\mathrm`）后面的
                #   空格是 **TeX 的定界符**，渲染时被吞掉，不是文字空格。
                #   `K=1.0\times 10^{2}` 必须落成 `K=1.0×10²`，不能落成 `K=1.0× 10²`。
                #   只对纯字母命令生效：`\,` / `\ ` / `\;` 这些**控制符号**后面
                #   的空格是有意义的，不能被吃掉。
                if m.group(1).isalpha():
                    while j < n and s[j] == " ":
                        j += 1
                i = j
                continue
            # —— 普通字符 ——
            emit(ch, vert)
            i += 1
        return True

    ok = walk(latex, None)
    if ok is None or ok is False:
        return None
    return out or None


# ------------------------------------------------ 免排版判据（简单等式 → 文字）
# 「量符号 = 数值 + 可选单位」这一种形态，没有任何二维排版需求。
# 判据只认**最窄**的形态，宁可留 OMML：多字母、链式等式、括号、\frac、
# 带幂次的单位（`a=9.8\,\mathrm{m/s^{2}}`）全部不在此列。
_NO_LAYOUT_LHS = re.compile(r"[A-Za-z]'?")             # `s` / `f'`
_NO_LAYOUT_VALUE = re.compile(r"([-+]?\d+(?:\.\d+)?)(.*)", re.S)
# `\text{...}` / `\mathrm{...}` 组：组内不许有花括号/反斜杠/角标 ——
# 有角标的（`\mathrm{m/s^{2}}`）本来就该交给混合策略，不该走这条免排版路。
_NO_LAYOUT_UNIT_GROUP = re.compile(
    r"(?:\\[ ,:;])?(?:\\text|\\mathrm|\\textrm|\\operatorname)\{([^{}\\^_]*)\}")

# 单位因子的分隔符：`/` 读作「每」，`·`/`⋅` 是连乘。
_UNIT_TAIL_SEP = re.compile(r"[\/\u00b7\u22c5]")

# 单字符后缀单位：`linear_to_latex.UNITS` 收的是「可独立成因子」的量纲符号，
# 而温度/角度这两个是**紧贴数值的后缀**，不在那张表里。不补的话 `T=25℃`、
# `α=30°` 会白留一个 OMML 公式对象 —— 保守方向出错（不误伤），但温度是
# 理科卷高频量，该降。
_UNIT_TAIL_EXTRA = frozenset({"\u2103", "\u00b0C", "\u00b0", "%"})


def _is_unit_tail(tail: str) -> bool:
    r"""*tail* 是不是一串货真价实的单位因子。

    要求每个因子都命中 `linear_to_latex.UNITS`（宁可逐字查表，也不放行
    「任意字母串」）：放开的话 `y=2x`、`n=2t` 会被读成「2 米」「2 吨」而降级，
    而那两条是**函数关系**，量符号的数学斜体不能丢。`t` 刻意不在 UNITS 里
    （`v=s/t` 必须保持真分式），这里因此也自动拒掉 `n=2t`。
    """
    factors = _UNIT_TAIL_SEP.split(tail)
    if not factors or any(not f for f in factors):
        return False                                     # `m/`、`/s`、`m//s`
    return all(f in L.UNITS or f in _UNIT_TAIL_EXTRA for f in factors)


def no_layout_reason(latex: str) -> str | None:
    r"""这个片段「不需要二维排版」的理由；需要 OMML 时返回 None。

    与 `is_pure_script` **互相独立**：那个函数是「真角标」的判据（要求带
    `^`/`_`），语义不同，混用会同时破坏两件事（HANDOFF-ROUND3 §11.4.5 ⚠️）。

    四条全部成立才降级：

    1. `structural_reason()` 为 None（没有分式/根号/积分/上下限/矩阵…）；
    2. 不含 `^` / `_`（没有角标，二维排版无从谈起）；
    3. 整串形如「单个量符号 = 数值 [+ 单位]」——
       LHS 一个 ASCII 字母（可带尾随 `'`），RHS 一个可带符号的十进制数，
       后面可选一个单位（裸写 `km` 或 `\\,\\text{...}` / `\\,\\mathrm{...}`，
       单位因子必须是 `linear_to_latex.UNITS` 的成员）；
    4. `_latex_to_pieces()` 能展开（不是未知命令）。

    刻意**不**降级：`F=ma`、`U=IR`、`I=U/R`、`v=s/t`（RHS 不是纯数值）、
    `f(x)=0`、`A=(-1,3)`（LHS 不是单字母）、`n\\geq 2`（没有 `=`）、
    `m=\\rho V=1.2\\times 60=72kg`（多段等式）、`a=9.8\\,\\mathrm{m/s^{2}}`（带角标）、
    `y=2x`（`x` 不是单位）。
    """
    if not latex:
        return None
    if "^" in latex or "_" in latex:
        return None
    if structural_reason(latex) is not None:
        return None
    parts = latex.split("=")
    if len(parts) != 2:                       # 链式等式 / 没有等号：保守留 OMML
        return None
    lhs, rhs = parts[0].strip(), parts[1].strip()
    if not _NO_LAYOUT_LHS.fullmatch(lhs):
        return None
    m = _NO_LAYOUT_VALUE.fullmatch(rhs)
    if not m:
        return None
    tail = m.group(2).strip()
    if tail:
        group = _NO_LAYOUT_UNIT_GROUP.fullmatch(tail)
        unit = group.group(1) if group else tail
        if not _is_unit_tail(unit.strip()):
            return None
    if _latex_to_pieces(latex) is None:
        return None
    return "plain_quantity" if tail else "plain_equation"


def _script_runs(latex: str, rpr):
    r"""纯上下标 LaTeX → 一串普通 ``w:r``，角标走 ``w:rPr/w:vertAlign``。

    `w:vertAlign` 的注入**复用** `docx_crossref.with_vert_align()`：它按
    `RPR_ORDER`（wml.xsd 的 CT_RPr 顺序）插入，写错位置的 run 属性会被 Word
    静默丢弃 —— 这是本轮唯一不能自己重写的一段逻辑，所以只读引用、不改那个文件。
    """
    from docx_crossref import with_vert_align          # 只读复用，避免重复实现

    pieces = _latex_to_pieces(latex)
    if not pieces:
        return None
    runs = []
    for text, vert in pieces:
        r = OxmlElement("w:r")
        new_rpr = with_vert_align(rpr, vert)
        if new_rpr is not None:
            r.append(new_rpr)
        t = OxmlElement("w:t")
        t.set(XMLSPACE, "preserve")
        t.text = text
        r.append(t)
        runs.append(r)
    return runs


# 段末独立编号：(1) / (2-1) / (3.2)
_EQNUM_RE = re.compile(r"[\s\t\u3000]*\((\d+(?:[-.]\d+)?)\)[\s\t\u3000]*$")


def _extract_eqnum(seg: dict) -> tuple[str, str] | None:
    """公式片段尾部若是独立编号 (n) → 剥出来，返回 (编号, 剥后公式原文)。

    剥掉编号后必须仍含关系符/强符号（否则不算「编号公式」，保守不动）。
    """
    orig = "".join(a["ch"] for a in seg["atoms"])
    m = _EQNUM_RE.search(orig)
    if not m:
        return None
    rest = orig[:m.start()].rstrip()
    if not (any(c in L.OP_CHARS for c in rest) or any(c in L.STRONG for c in rest)):
        return None
    tex = re.sub(r"\s*\(\d+(?:[-.]\d+)?\)$", "", seg["latex"]).strip()
    return m.group(1), tex


# --------------------------------------------------- 交叉引用 / 多行对齐编号
_TAG_RE = re.compile(r"\\tag\*?\{([^{}]*)\}")
_LABEL_RE = re.compile(r"\\label\{([^{}]*)\}")
_REF_CMD_RE = re.compile(r"\\(eqref|ref|label|tag)\{([^{}]*)\}")
# Anchored at the environment, but tolerant of a leading fragment such as
# "A = " before \begin{align}: the caller passes a whole formula, and the
# environment is often only the tail of it.
_ALIGN_RE = re.compile(
    r"\\begin\{(align|alignat|eqnarray|gather)(\*?)\}(.*)\\end\{\1\2\}", re.S)


def _env_ranges(text: str) -> list[tuple[int, int]]:
    r"""所有 \begin{…}…\end{…} 的字符区间（按嵌套深度配对）。"""
    out: list[tuple[int, int]] = []
    i = 0
    while True:
        s = text.find("\\begin{", i)
        if s < 0:
            break
        depth, j = 0, s
        while j < len(text):
            b = text.find("\\begin{", j)
            e = text.find("\\end{", j)
            if e < 0:
                break
            if 0 <= b < e:
                depth += 1
                j = b + 7
            else:
                depth -= 1
                j = e + 5
                if depth == 0:
                    out.append((s, j))
                    break
        i = max(s + 1, j)
    return out


def _mark_commands(atoms: list[dict]) -> list[dict]:
    r"""把正文里的 \eqref{} / \ref{} / \label{} / \tag{} 摘成 special atom。

    环境内部的 \tag / \label 原样保留（由 _parse_align 在行级消费）。
    special atom 的 ch 是空串，不参与数学片段判定。
    """
    text = "".join(a["ch"] for a in atoms)
    envs = _env_ranges(text)

    def in_env(pos: int) -> bool:
        return any(s <= pos < e for s, e in envs)

    out: list[dict] = []
    i = 0
    while i < len(atoms):
        m = _REF_CMD_RE.match(text[i:])
        if m and not in_env(i):
            out.append({"ch": "", "rpr": atoms[i]["rpr"], "script": None,
                        "special": (m.group(1), m.group(2).strip())})
            i += m.end()
            continue
        out.append(atoms[i])
        i += 1
    return out


def _split_env_rows(body: str) -> list[str]:
    r"""按顶层 \\ 切行（嵌套环境里的 \\ 不切）。"""
    rows, cur = [], []
    depth, i = 0, 0
    while i < len(body):
        if body.startswith("\\begin{", i):
            depth += 1
            cur.append("\\begin{")
            i += 7
            continue
        if body.startswith("\\end{", i):
            depth -= 1
            cur.append("\\end{")
            i += 5
            continue
        if body.startswith("\\\\", i) and depth == 0:
            rows.append("".join(cur))
            cur = []
            i += 2
            continue
        cur.append(body[i])
        i += 1
    if "".join(cur).strip():
        rows.append("".join(cur))
    return rows


def _parse_align(latex: str) -> list[dict] | None:
    r"""`\begin{align} a &= b \tag{1} \\ c &= d \end{align}` → 行结构。

    ``search``, not ``match``: a formula often carries a leading fragment such
    as ``A = `` before the environment, and only the environment itself is of
    interest here.
    """
    m = _ALIGN_RE.search(latex)
    if not m:
        return None
    rows: list[dict] = []
    for raw in _split_env_rows(m.group(3)):
        tag = label = None
        tm = _TAG_RE.search(raw)
        if tm:
            tag = tm.group(1).strip()
            raw = raw[:tm.start()] + raw[tm.end():]
        lm = _LABEL_RE.search(raw)
        if lm:
            label = lm.group(1).strip()
            raw = raw[:lm.start()] + raw[lm.end():]
        if not raw.strip():
            continue
        if "&" in raw:
            left, right = raw.split("&", 1)
            right = right.replace("&", "\\ ")      # 多对齐点降级成空格
        else:
            left, right = raw, ""
        rows.append({"left": left.strip(), "right": right.strip(),
                     "tag": tag, "label": label})
    return rows or None


def _bookmark_name(key: str) -> str:
    safe = re.sub(r"[^0-9A-Za-z_]", "_", key)
    if not safe or safe[0].isdigit():
        safe = "e" + safe
    return ("eq_" + safe)[:40]


def _bm_id(stats: dict) -> int:
    stats["bm"] = stats.get("bm", 9000) + 1
    return stats["bm"]


def _bookmark(bm_id: int, name: str, end: bool = False):
    el = OxmlElement("w:bookmarkEnd" if end else "w:bookmarkStart")
    if not end:
        el.set(qn("w:name"), name)
    el.set(qn("w:id"), str(bm_id))
    return el


def _ref_runs(key: str, number: str | None, rpr) -> list:
    """REF 域：显示编号文本，Word 里按 F9 依书签内容更新。"""
    name = _bookmark_name(key)

    def fld(kind: str):
        r = _run_tag()
        f = OxmlElement("w:fldChar")
        f.set(qn("w:fldCharType"), kind)
        r.append(f)
        return r

    ins = _run_tag()
    it = OxmlElement("w:instrText")
    it.set(XMLSPACE, "preserve")
    it.text = f" REF {name} \\h "
    ins.append(it)
    shown = (number if number and number.startswith("(")
             else (f"({number})" if number else "(?)"))
    return [fld("begin"), ins, fld("separate"),
            _text_run(rpr, shown), fld("end")]

def _make_tc(width: int, jc: str, els: list):
    tc = OxmlElement("w:tc")
    tcpr = OxmlElement("w:tcPr")
    w = OxmlElement("w:tcW")
    w.set(qn("w:w"), str(width))
    w.set(qn("w:type"), "dxa")
    tcpr.append(w)
    va = OxmlElement("w:vAlign")
    va.set(qn("w:val"), "center")
    tcpr.append(va)
    tc.append(tcpr)
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    j = OxmlElement("w:jc")
    j.set(qn("w:val"), jc)
    ppr.append(j)
    p.append(ppr)
    for el in els:
        p.append(el)
    # 单元格里公式独占段落时 Word 会提升成行间（居中、∑ 纵向堆叠），
    # 用 U+200B 压住以保留「单元格内行内 + 列界对齐」的行为。
    r = _run_tag()
    t = _t_tag()
    t.text = "\u200b"
    r.append(t)
    p.append(r)
    tc.append(p)
    return tc


def _number_els(tag: str, label: str | None, stats: dict) -> list:
    """编号内容 (n)，有 label 则整段包书签（供 REF 域引用）。"""
    shown = tag if tag.startswith("(") else f"({tag})"
    els: list = []
    bm = None
    if label:
        bm = _bm_id(stats)
        els.append(_bookmark(bm, _bookmark_name(label)))
    els.append(_text_run(None, shown))
    if bm is not None:
        els.append(_bookmark(bm, _bookmark_name(label), end=True))
    return els


def _om_or_scripts(tex: str, rpr, stats: dict) -> list | None:
    r"""混合策略分流：纯上下标 / 免排版简单等式 → 文字 run；有结构 → OMML。

    返回 ``None`` 表示「按原路走 OMML」（调用方继续用它既有的错误处理）。

    两层判据的先后不是随意的：
      * 免排版层（`no_layout_reason`）只管「压根不需要公式对象」的简单等式；
      * 文字角标层（`is_pure_script`）只管「需要角标、但不需要二维排版」的。
    第二层要求片段带 `^`/`_`，第一层恰恰**要求不带**，两者互斥、谁也不改谁。
    """
    if not MIXED_SUPERSCRIPT:
        # MIXED_SUPERSCRIPT 是「全部 OMML」的总闸：它关了，免排版也一起停，
        # 否则「置 False 就等于旧行为」这句承诺就不成立。
        return None
    if NO_LAYOUT_NEEDED:
        reason = no_layout_reason(tex)
        if reason is not None:
            runs = _script_runs(tex, rpr)
            if runs:
                stats["plain_no_layout"] = stats.get("plain_no_layout", 0) + 1
                stats.setdefault("plain_no_layout_rows", []).append(
                    {"latex": tex, "reason": reason})
                return runs
    if not is_pure_script(tex):
        return None
    runs = _script_runs(tex, rpr)
    if not runs:
        return None
    stats["text_scripts"] = stats.get("text_scripts", 0) + 1
    stats.setdefault("text_script_rows", []).append(tex)
    return runs


def _omml_with_phasor(tex: str, stats: dict | None = None):
    r"""LaTeX→OMML 的统一出口：转换后顺带修相量点（R-1）。

    latex2mathml 把 ``\dot{X}`` 转成 ``m:limUpp``（上限），Word 里相量点会
    浮在字上方且过大；本包装器在转换后立即调 ``fix_phasor_dots`` 把它就地
    改写成 accent（``m:acc`` + U+0307）。本文件**全部** LaTeX→OMML 转换都
    必须从这里过（历史坑：只在单点挂修复会漏 ``rebuild()`` 的分支，
    见 MERGE-MAP R-1）。

    异常语义（RT-001 修法的关键一条）
    ----------------------------------
    两级异常**必须分开**：

    1. ``latex_to_omml_lxml`` 失败 → 原样抛出，由各调用点既有的 try/except
       处理，记 ``stats["failed"]`` 并回退纯文本。这是「公式没转成」。
    2. ``fix_phasor_dots`` 失败 → **只在内部捕获**，记进
       ``stats["phasor_errors"]``，并返回**未修但可用**的 OMML。
       这是「公式转成了、相量点没修好」，不该被当成转换失败——否则一篇
       文档里一个 ``\\dot`` 出问题就会让整段公式回退成纯文本，代价远大于收益。

    ``stats`` 非 None 时，修复成功累进 ``stats["phasor_fixes"]``（汇总行、
    ``--report`` 与测试都读它），失败追加 ``{"latex", "error"}`` 进
    ``stats["phasor_errors"]``（``error`` 里带异常类型名，便于区分根因）。
    """
    om = latex_to_omml_lxml(tex)
    try:
        n = fix_phasor_dots(om)
    except Exception as exc:              # 修相量点出事不该拖垮整篇转换
        if stats is not None:
            stats.setdefault("phasor_errors", []).append(
                {"latex": tex,
                 "error": f"{type(exc).__name__}: {str(exc)[:160]}"[:200]})
        return om                         # 返回未修版本：公式本身仍然可用
    if n and stats is not None:
        stats["phasor_fixes"] = stats.get("phasor_fixes", 0) + n
    return om


def _om_or_text(tex: str, fallback: str, stats: dict) -> list:
    if not tex:
        return []
    runs = _om_or_scripts(tex, None, stats)
    if runs is not None:
        return runs
    try:
        return [_omml_with_phasor(tex, stats)]
    except Exception as exc:
        stats["failed"].append({"latex": tex, "error": str(exc)[:200]})
        return [_text_run(None, fallback)]


def _build_align_table(rows: list[dict], total_tw: int, stats: dict):
    """多行对齐公式 + 逐行编号 → 无边框 3 列表格。

    左部右对齐、右部左对齐（`&` 对齐点落在列界，还原 align 语义），
    编号列固定宽右对齐。这是 Word 里唯一能同时保证「对齐 + 逐行右缘编号」的结构。
    """
    num_w = 1100
    body_w = max(total_tw - num_w, 2400)
    left_w = body_w // 3
    right_w = body_w - left_w

    tbl = OxmlElement("w:tbl")
    tblpr = OxmlElement("w:tblPr")
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), str(total_tw))
    tw.set(qn("w:type"), "dxa")
    tblpr.append(tw)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    tblpr.append(jc)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        b = OxmlElement(f"w:{side}")
        b.set(qn("w:val"), "none")
        b.set(qn("w:sz"), "0")
        b.set(qn("w:space"), "0")
        borders.append(b)
    tblpr.append(borders)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tblpr.append(layout)
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        m = OxmlElement(f"w:{side}")
        m.set(qn("w:w"), "0")
        m.set(qn("w:type"), "dxa")
        mar.append(m)
    tblpr.append(mar)
    tbl.append(tblpr)

    grid = OxmlElement("w:tblGrid")
    for w in (left_w, right_w, num_w):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    tbl.append(grid)

    for r in rows:
        tr = OxmlElement("w:tr")
        trpr = OxmlElement("w:trPr")
        trpr.append(OxmlElement("w:cantSplit"))
        tr.append(trpr)
        tr.append(_make_tc(left_w, "right", _om_or_text(r["left"], r["left"], stats)))
        tr.append(_make_tc(right_w, "left", _om_or_text(r["right"], r["right"], stats)))
        num_els = _number_els(r["tag"], r["label"], stats) if r["tag"] else []
        tr.append(_make_tc(num_w, "right", num_els))
        tbl.append(tr)
    return tbl


def _eqnum_paragraph(tex: str, num: str, label: str | None,
                     mid: int, right: int, stats: dict):
    """单行编号公式段（gather 逐行拆段用；与 rebuild 编号分支同构）。"""
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    tabs = OxmlElement("w:tabs")
    for kind, pos in (("center", mid), ("right", right)):
        t = OxmlElement("w:tab")
        t.set(qn("w:val"), kind)
        t.set(qn("w:pos"), str(pos))
        tabs.append(t)
    ppr.append(tabs)
    j = OxmlElement("w:jc")
    j.set(qn("w:val"), "left")
    ppr.append(j)
    p.append(ppr)
    r = _run_tag()
    r.append(_tab_tag())
    p.append(r)
    for el in _om_or_text(tex, tex, stats):
        p.append(el)
    r = _run_tag()
    r.append(_tab_tag())
    p.append(r)
    for el in _number_els(num, label, stats):
        p.append(el)
    return p


def _run_tag():
    return OxmlElement("w:r")


def _tab_tag():
    return OxmlElement("w:tab")


def _t_tag():
    return OxmlElement("w:t")


def _text_run(rpr, text: str):
    """按原 rPr 造一个 w:r；文本里的 \\t 展开为 w:tab。"""
    r = _run_tag()
    if rpr is not None:
        r.append(deepcopy(rpr))
    for i, part in enumerate(text.split("\t")):
        if i:
            r.append(_tab_tag())
        if part:
            t = _t_tag()
            t.text = part
            if part != part.strip():
                t.set(XMLSPACE, "preserve")
            r.append(t)
    return r


def rebuild(p_el, segs: list[dict], stats: dict,
            eqnum: dict | None = None, label_map: dict | None = None) -> None:
    ppr = p_el.find(qn("w:pPr"))
    for child in list(p_el):
        if child is not ppr:
            p_el.remove(child)

    if eqnum is not None:
        # 编号公式段：tab → 公式 → tab → (n)；tabs/jc 已由调用方写进 pPr。
        # Word 桌面版不渲染 #(n) 语法（2026-10-24 实测，见 latex-fix），
        # 行业标准做法就是「居中制表位 + 右对齐制表位」双 tab。
        text_runs = _om_or_scripts(eqnum["latex"], eqnum.get("rpr"), stats)
        if text_runs is not None:
            r = _run_tag()
            r.append(_tab_tag())
            p_el.append(r)
            for el in text_runs:
                p_el.append(el)
            r = _run_tag()
            r.append(_tab_tag())
            p_el.append(r)
            for el in _number_els(eqnum["num"], eqnum.get("label"), stats):
                p_el.append(el)
            return
        om = None
        try:
            om = _omml_with_phasor(eqnum["latex"], stats)
        except Exception as exc:                       # 失败 → 退回普通段落
            stats["failed"].append({"latex": eqnum["latex"], "error": str(exc)[:200]})
        if om is not None:
            r = _run_tag()
            r.append(_tab_tag())
            p_el.append(r)
            p_el.append(om)
            r = _run_tag()
            r.append(_tab_tag())
            p_el.append(r)
            for el in _number_els(eqnum["num"], eqnum.get("label"), stats):
                p_el.append(el)
            return
        # 走到这 = 公式转换失败，按原文重建（编号留在文本里）
        p_el.append(_text_run(eqnum["rpr"], eqnum["orig"]))
        return

    n_math = 0
    for seg in segs:
        if seg["kind"] == "text":
            buf: list[dict] = []

            def flush():
                if buf:
                    p_el.append(_text_run(buf[0]["rpr"],
                                          "".join(a["ch"] for a in buf)))
                    buf.clear()

            for a in seg["atoms"]:
                sp = a.get("special")
                if not sp:
                    buf.append(a)
                    continue
                flush()
                cmd, key = sp
                if cmd in ("eqref", "ref"):
                    num = (label_map or {}).get(key)
                    if num is None:
                        # 未定义的引用：不建 REF 域（Word 会显示「错误!未找到引用源」），
                        # 直接落 (?) 并在报告里提示补 \label
                        p_el.append(_text_run(a["rpr"], "(?)"))
                        stats.setdefault("bad_refs", []).append(key)
                    else:
                        stats["refs"] = stats.get("refs", 0) + 1
                        for el in _ref_runs(key, num, a["rpr"]):
                            p_el.append(el)
                # label / tag 不渲染（编号处已放书签）
            flush()
            continue
        text_runs = _om_or_scripts(seg["latex"], (seg["atoms"][0]["rpr"]
                                                 if seg["atoms"] else None), stats)
        if text_runs is not None:
            for el in text_runs:
                p_el.append(el)
            continue
        try:
            om = _omml_with_phasor(seg["latex"], stats)
        except Exception as exc:                       # 失败 → 退回原文
            stats["failed"].append({"latex": seg["latex"], "error": str(exc)[:200]})
            p_el.append(_text_run(seg["atoms"][0]["rpr"],
                                  "".join(a["ch"] for a in seg["atoms"])))
            continue
        p_el.append(om)
        n_math += 1

    # 整段只有公式 → 补 U+200B，防止 Word 当行间公式渲染
    if n_math and not any(s["kind"] == "text" and "".join(a["ch"] for a in s["atoms"]).strip()
                          for s in segs):
        r = _run_tag()
        t = _t_tag()
        t.text = "\u200b"
        r.append(t)
        p_el.append(r)


def _setup_eqnum_para(p_el, mid: int, right: int) -> None:
    """编号公式段：左对齐 + 居中/右对齐双制表位（编号用）。"""
    pf = Paragraph(p_el, None).paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pf.tab_stops.add_tab_stop(Twips(mid), WD_TAB_ALIGNMENT.CENTER)
    pf.tab_stops.add_tab_stop(Twips(right), WD_TAB_ALIGNMENT.RIGHT)


def _analyze(atoms: list[dict], fix_merged: bool, stats: dict | None = None,
             idx: int = 0):
    """段落 → (segs, maths, eqnum, labels, align_rows)。

    编号段 / 对齐环境 / \\label 的判定都走这里：第一遍（收集 label）与第二遍
    （重建）共用同一套判定，保证 label → 编号映射一致。
    """
    segs = L.split_paragraph(atoms)
    maths = [s for s in segs if s["kind"] == "math"]
    if fix_merged:
        for s in maths:
            before = s["latex"]
            s["latex"], fixes = L.repair_merged_scripts(before)
            if fixes and stats is not None:
                stats["script_fixes"].append(
                    {"p": idx, "before": before, "after": s["latex"], "pairs": fixes})

    eqnum = None
    if len(maths) == 1:
        last = segs[-1]
        tail_blank = (last["kind"] == "text"
                      and not "".join(a["ch"] for a in last["atoms"]).strip())
        if last is maths[0] or tail_blank:
            got = _extract_eqnum(maths[0])
            if got:
                eqnum = {"num": got[0], "latex": got[1],
                         "rpr": maths[0]["atoms"][0]["rpr"],
                         "orig": "".join(a["ch"] for a in atoms)}
                maths[0]["latex"] = got[1]
        if eqnum is None and len(segs) >= 2 and segs[-1]["kind"] == "text":
            m2 = _EQNUM_RE.fullmatch("".join(a["ch"] for a in segs[-1]["atoms"]))
            if m2:
                eqnum = {"num": m2.group(1), "latex": maths[0]["latex"],
                         "rpr": segs[-1]["atoms"][0]["rpr"],
                         "orig": "".join(a["ch"] for a in atoms)}
                segs = segs[:-1]

    labels: list[str] = []
    for seg in segs:
        if seg["kind"] == "text":
            for a in seg["atoms"]:
                sp = a.get("special")
                if sp and sp[0] == "label":
                    labels.append(sp[1])
    if eqnum is not None and labels:
        eqnum["label"] = labels[0]

    align_rows = _parse_align(maths[0]["latex"]) if len(maths) == 1 else None
    return segs, maths, eqnum, labels, align_rows


def mathify(src: Path, out: Path | None, report: Path | None, tex: Path | None,
            fix_merged: bool = True, dry_run: bool = False,
            number_align: bool = False) -> dict:
    doc = Document(str(src))
    body = doc.element.body

    stats = {"paragraphs": 0, "math_segments": 0, "failed": [],
             "script_fixes": [], "eqnum": 0, "align_tables": 0,
             "refs": 0, "unit_strings": 0, "bm": 9000,
             "text_scripts": 0, "text_script_rows": [],
             "plain_no_layout": 0, "plain_no_layout_rows": [],
             # 相量点修复：成功计数 / 失败明细（RT-001 补的两个键）。
             # 失败项刻意**不**并进 "failed"——那是「公式没转成」的语义，
             # 而相量点修不好时公式本身是好的，只影响观感。
             "phasor_fixes": 0, "phasor_errors": []}
    rows: list[dict] = []

    # 版心宽度（twips）：编号公式段的居中/右对齐制表位
    sec = doc.sections[0]
    right_pos = Emu(int(sec.page_width) - int(sec.left_margin)
                    - int(sec.right_margin)).twips
    mid_pos = right_pos // 2

    paras = list(body.iter(qn("w:p")))

    # ---- 第一遍：收集 \label → 编号（前向引用要能取到） ----
    label_map: dict[str, str] = {}
    for p_el in paras:
        atoms = L.paragraph_atoms(p_el)
        if not atoms:
            continue
        _s, _m, eqnum1, labels1, align1 = _analyze(
            _mark_commands(atoms), fix_merged=False)
        if eqnum1 is not None and labels1:
            for k in labels1:
                label_map[k] = eqnum1["num"]
        for r in (align1 or []):
            if r.get("label") and r.get("tag"):
                label_map[r["label"]] = r["tag"]

    # ---- 第二遍：重建 ----
    for idx, p_el in enumerate(paras):
        atoms = L.paragraph_atoms(p_el)
        if atoms is None or not atoms:
            continue
        orig = "".join(a["ch"] for a in atoms)
        segs, maths, eqnum, labels, align_rows = _analyze(
            _mark_commands(atoms), fix_merged, stats, idx)
        has_special = any(a.get("special") for s in segs if s["kind"] == "text"
                          for a in s["atoms"])
        if not maths and not has_special:
            continue
        stats["paragraphs"] += 1

        # —— 多行对齐环境 + 行编号 → 3 列表格（有 &）/ 逐行编号段落（无 &）——
        if align_rows is not None and (number_align or any(r["tag"] for r in align_rows)):
            if number_align:
                seq = stats.get("align_seq", 0)
                for r in align_rows:
                    if not r["tag"]:
                        seq += 1
                        r["tag"] = str(seq)
                stats["align_seq"] = seq
                for r in align_rows:
                    if r.get("label") and r.get("tag"):
                        label_map.setdefault(r["label"], r["tag"])
            if not dry_run:
                if any(r["right"] for r in align_rows):
                    p_el.addnext(_build_align_table(align_rows, right_pos, stats))
                    p_el.getparent().remove(p_el)
                else:
                    prev = p_el
                    for r in align_rows:
                        if not r["tag"]:
                            continue
                        np = _eqnum_paragraph(r["left"], r["tag"], r["label"],
                                              mid_pos, right_pos, stats)
                        prev.addnext(np)
                        prev = np
                    p_el.getparent().remove(p_el)
            stats["align_tables"] += 1
            for r in align_rows:
                body_tex = r["left"] + (" & " + r["right"] if r["right"] else "")
                stats["math_segments"] += 1
                units = L.unit_slash_chains(body_tex)
                stats["unit_strings"] += len(units)
                rows.append({"p": idx, "linear": body_tex.strip(),
                             "latex": body_tex.strip(), "context": orig[:120],
                             "units": units,
                             "num": f"({r['tag']})" if r["tag"] else ""})
            continue

        if not dry_run:
            if eqnum is not None:
                _setup_eqnum_para(p_el, mid_pos, right_pos)
            rebuild(p_el, segs, stats, eqnum=eqnum, label_map=label_map)

        if eqnum is not None:
            stats["eqnum"] += 1
        for s in maths:
            linear = "".join(a["ch"] for a in s["atoms"])
            units = L.unit_slash_chains(linear)
            stats["math_segments"] += 1
            stats["unit_strings"] += len(units)
            rows.append({
                "p": idx,
                "linear": linear,
                "latex": s["latex"],
                "context": orig[:120],
                "units": units,
                "num": f"({eqnum['num']})" if eqnum is not None else "",
                # 报告要能按段落号核对降级明细，所以理由在这里复算一遍：
                # 两个开关都开着才真的会降级，否则报告不能谎报。
                "no_layout": (no_layout_reason(s["latex"])
                              if NO_LAYOUT_NEEDED and MIXED_SUPERSCRIPT else None),
            })

    if not dry_run and out is not None:
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))

    if report is not None:
        # dry-run 时 rebuild 不跑，`stats['plain_no_layout']` 还是 0；但报告要能预览
        # 「哪些会被降级」，所以这种情况下用明细条数顶上，别让摘要和表打架。
        plain = [r for r in rows if r.get("no_layout")]
        n_plain = len(plain) if dry_run else stats.get("plain_no_layout", 0)
        lines = ["# 公式 → LaTeX 转换报告", "",
                 f"- 源文件：`{src}`",
                 f"- 输出：`{out}`" if out else "- 输出：（dry-run，未写文件）",
                 f"- 含公式段落：{stats['paragraphs']}　转换片段：{stats['math_segments']}　"
                 f"编号公式段：{stats.get('eqnum', 0)}　对齐编号块：{stats.get('align_tables', 0)}　"
                 f"引用域：{stats.get('refs', 0)}　"
                 f"量纲串：{stats.get('unit_strings', 0)}　"
                 f"免排版降级：{n_plain}　"
                 f"修正相量点：{stats.get('phasor_fixes', 0)}　"
                 f"失败：{len(stats['failed'])}", "",
                 "| # | 段落 | 原线性写法 | 生成的 LaTeX | 编号 | 量纲串（`/` = 每） |",
                 "|---|---|---|---|---|---|"]
        for i, r in enumerate(rows, 1):
            lin = r["linear"].replace("|", "\\|")
            lat = r["latex"].replace("|", "\\|")
            uni = "、".join(r.get("units") or []).replace("|", "\\|")
            lines.append(f"| {i} | {r['p']} | `{lin}` "
                         f"| `{lat}` | {r.get('num', '')} | {uni} |")
        if stats["script_fixes"]:
            lines += ["", "## 已自动修正的源排版错误（上下标把后面的正常字符一起包了进去）", "",
                      "| 段落 | 修正前 | 修正后 |", "|---|---|---|"]
            for f in stats["script_fixes"]:
                lines.append(f"| {f['p']} | `{f['before']}` | `{f['after']}` |")
        unit_rows = [(r["p"], u) for r in rows for u in (r.get("units") or [])]
        if unit_rows:
            lines += ["", "## 识别为量纲串的 `/`（读作「每」，未转分式）", ""]
            for p, u in unit_rows:
                lines.append(f"- 段落 {p}：`{u}`")
        plain = [r for r in rows if r.get("no_layout")]
        if plain:
            lines += ["", "## 免排版降级（量符号=数值[+单位]，已落成普通文字 run，未产 OMML）", "",
                      "| 段落 | 原线性写法 | 生成的 LaTeX | 理由 |", "|---|---|---|---|"]
            for r in plain:
                lin2 = r["linear"].replace("|", "\\|")
                lat2 = r["latex"].replace("|", "\\|")
                lines.append(f"| {r['p']} | `{lin2}` | `{lat2}` | {r['no_layout']} |")
        left = [r for r in rows if re.search(r"[_^]\{(\d+[A-Z][A-Za-z0-9]*)\}", r["latex"])]
        if left:
            lines += ["", "## ⚠️ 仍需人工确认的上下标", ""]
            for r in left:
                lines.append(f"- 段落 {r['p']}：`{r['latex']}`")
        if stats["failed"]:
            lines += ["", "## 转换失败（已保留原文）", ""]
            for f in stats["failed"]:
                lines.append(f"- `{f['latex']}` → {f['error']}")
        # 相量点修复失败**不是**转换失败（公式已转成、只是 ``\dot`` 可能仍是
        # limUpp 形态），所以单列一节，别并进上面那段——两件事混在一起，
        # 用户会以为公式没转成。
        perr = stats.get("phasor_errors") or []
        if perr:
            lines += ["", "## ⚠️ 相量点未修复（公式已转换，`\\dot` 可能仍是 limUpp 形态）", ""]
            for f in perr:
                lines.append(f"- `{f['latex']}` → {f['error']}")
        bad = stats.get("bad_refs") or []
        if bad:
            lines += ["", "## ⚠️ 未解析的引用（源文档里没有对应 \\label，已落成 (?)）", ""]
            for k in dict.fromkeys(bad):
                lines.append(f"- `\\eqref{{{k}}}` → 缺少 `\\label{{{k}}}`")
        report.parent.mkdir(parents=True, exist_ok=True)
        report.write_text("\n".join(lines), encoding="utf-8")

    if tex is not None:
        seen, out_lines = set(), []
        for r in rows:
            if r["latex"] in seen:
                continue
            seen.add(r["latex"])
            out_lines.append(f"% 段落 {r['p']}：{r['context'][:80]}")
            out_lines.append(f"\\[{r['latex']}\\]")
            out_lines.append("")
        tex.parent.mkdir(parents=True, exist_ok=True)
        tex.write_text("\n".join(out_lines), encoding="utf-8")

    return stats


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(
        description="把 docx 里的线性公式批量换成 LaTeX 渲染的原生 OMML。"
    )
    p.add_argument("input", help="源 .docx（只读）")
    p.add_argument("-o", "--out", help="输出 .docx；省略则需配 --inplace 或 --dry-run")
    p.add_argument("--inplace", action="store_true",
                   help="就地覆盖（会先备份成 <名>.bak.docx）")
    p.add_argument("--report", help="写出 Markdown 报告（原写法 ↔ LaTeX）")
    p.add_argument("--tex", help="写出 LaTeX 清单（去重）")
    p.add_argument("--no-merged-script-fix", action="store_true",
                   help="不自动拆开源文档里合并过头的上下标")
    p.add_argument("--dry-run", action="store_true", help="只出报告，不改文档")
    p.add_argument("--number-align", action="store_true",
                   help="align/gather 里没有 \\tag 的行按顺序自动编号（默认只认 \\tag{...}）")
    args = p.parse_args(argv)

    src = Path(args.input)
    if not src.is_file():
        print(f"error: {src} 不存在", file=sys.stderr)
        return 1

    if args.inplace and not args.dry_run:
        bak = src.with_suffix(".bak.docx")
        shutil.copy2(src, bak)
        out = src
        print(f"备份: {bak}")
    elif args.out:
        out = Path(args.out)
    elif args.dry_run:
        out = None
    else:
        print("error: 需要 -o/--out、--inplace 或 --dry-run 之一", file=sys.stderr)
        return 2

    stats = mathify(src, out, Path(args.report) if args.report else None,
                    Path(args.tex) if args.tex else None,
                    fix_merged=not args.no_merged_script_fix,
                    dry_run=args.dry_run,
                    number_align=args.number_align)

    print(f"段落 {stats['paragraphs']}　片段 {stats['math_segments']}　"
          f"编号段 {stats.get('eqnum', 0)}　对齐编号块 {stats.get('align_tables', 0)}　"
          f"引用域 {stats.get('refs', 0)}　量纲串 {stats.get('unit_strings', 0)}　"
          f"失败 {len(stats['failed'])}　"
          f"修正上下标 {len(stats['script_fixes'])}　"
          f"修正相量点 {stats.get('phasor_fixes', 0)}")
    if out:
        print(f"输出: {out}")
    for f in stats["failed"][:10]:
        print("  FAIL", f["latex"], "->", f["error"][:120])
    # 相量点没修好**不算转换失败**（公式本身是好的），所以只提示、不改退出码。
    # 走 stderr：本脚本的 stdout 是给人看的汇总行，诊断一律走另一路。
    perr = stats.get("phasor_errors") or []
    for f in perr[:10]:
        print("  PHASOR-WARN", f["latex"], "->", f["error"], file=sys.stderr)
    if stats["failed"]:
        # Some formulas could not be converted. The document is still written
        # (partial conversion is useful), but the caller must be able to tell:
        # a silent 0 would hide broken equations in a delivered file.
        print(f"error: {len(stats['failed'])} formula(s) failed to convert",
              file=sys.stderr)
        return 3
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
