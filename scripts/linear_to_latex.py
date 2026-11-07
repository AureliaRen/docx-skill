# -*- coding: utf-8 -*-
"""线性公式 → LaTeX。

处理对象：文档里「Unicode 字符 + 真上下标 run」写成的公式
（`U`+下标`oc`、`P∝I`+上标`2`、`√2`、`∠53.1°`、`1kΩ=10³Ω`、`f′(x)` …）。

入口：
    find_math_spans(atoms)      哪些连续片段算公式（宁缺毋滥）
    atoms_to_linear(...)        原子序列 → 带 ^{} _{} 的线性串
    linear_to_latex(linear)     线性串 → LaTeX
    repair_merged_scripts(tex)  修源文档里「把后续字符一起设成上下标」的错误

约定：判据不满足的片段原样保留，绝不硬转。
"""
from __future__ import annotations

# --- repo bootstrap: make sibling imports work from any cwd ---------------
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix as _encoding_fix  # noqa: F401 — UTF-8 stdio on Windows
# --------------------------------------------------------------------------

import re

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# ---------------------------------------------------------------- 字符映射
GREEK = {
    "α": r"\alpha", "β": r"\beta", "γ": r"\gamma", "δ": r"\delta",
    "ε": r"\varepsilon", "ζ": r"\zeta", "η": r"\eta", "θ": r"\theta",
    "ι": r"\iota", "κ": r"\kappa", "λ": r"\lambda", "μ": r"\mu",
    "ν": r"\nu", "ξ": r"\xi", "π": r"\pi", "ρ": r"\rho",
    "σ": r"\sigma", "τ": r"\tau", "υ": r"\upsilon", "φ": r"\varphi",
    "χ": r"\chi", "ψ": r"\psi", "ω": r"\omega",
    "Γ": r"\Gamma", "Θ": r"\Theta", "Λ": r"\Lambda", "Ξ": r"\Xi",
    "Π": r"\Pi", "Υ": r"\Upsilon", "Φ": r"\Phi", "Ψ": r"\Psi",
}

# 显式写码位，避免编辑器把 U+00B7 / U+22C5 / U+2219 混成一个字形
SYMBOL = {
    "\u2212": "-", "\u2013": "-", "\u2014": "-", "\u2010": "-", "\u2011": "-",
    "\u00d7": r"\times", "\u00f7": r"\div",
    "\u00b7": r"\cdot", "\u22c5": r"\cdot", "\u2219": r"\cdot", "\u2217": r"\ast",
    "\u2248": r"\approx", "\u2264": r"\leq", "\u2265": r"\geq",
    "\u2260": r"\neq", "\u2261": r"\equiv", "\u221d": r"\propto",
    "\u00b1": r"\pm", "\u2213": r"\mp", "\u221e": r"\infty",
    "\u2220": r"\angle", "\u2211": r"\sum", "\u03a3": r"\sum",
    "\u222b": r"\int", "\u222e": r"\oint", "\u222c": r"\iint",
    "\u222d": r"\iiint", "\u2202": r"\partial", "\u2207": r"\nabla",
    "\u221a": r"\surd",
    "\u2192": r"\to", "\u2190": r"\leftarrow", "\u21d2": r"\Rightarrow",
    "\u2200": r"\forall", "\u2208": r"\in", "\u2245": r"\cong",
    "\u2282": r"\subset", "\u2283": r"\supset", "\u2286": r"\subseteq",
    "\u2229": r"\cap", "\u222a": r"\cup", "\u2295": r"\oplus",
    "\u2297": r"\otimes", "\u22a5": r"\perp", "\u00ac": r"\neg",
    "\u0394": r"\Delta", "\u03a9": r"\Omega",
    "\u2032": "'", "\u2033": "''", "\u2034": "'''", "\u02dc": r"\sim",
    "\u2225": r"\|", "\u2016": r"\|", "\u00b0": r"^{\circ}",
    "\u00bd": r"\frac{1}{2}", "\u00bc": r"\frac{1}{4}", "\u00be": r"\frac{3}{4}",
    "\u2153": r"\frac{1}{3}", "\u2154": r"\frac{2}{3}",
    "\u2009": r"\,", "\u2005": r"\,", "\u00a0": "~",
}
COMBINING = {"\u0307": "dot", "\u0308": "ddot", "\u0304": "bar",
             "\u0305": "bar", "\u20d7": "vec", "\u0302": "hat"}

# 关系/运算：片段里出现这些才考虑转换（箭头不算 —— 手册里它多是正文分隔符）
OP_CHARS = set("=\u2248\u2264\u2265\u2260\u2261\u221d\u00b1\u2213<>")
STRONG = set("\u221a\u2211\u222b\u2202\u03a3\u2207\u222e\u222c\u222d")
FUNCS = {"arcsin", "arccos", "arctan", "sinh", "cosh", "tanh",
         "sin", "cos", "tan", "cot", "sec", "csc",
         "lim", "max", "min", "sup", "inf", "det", "dim", "ker", "gcd",
         "ln", "lg", "log", "exp", "arg"}
# LaTeX 没有内置命令的算子名 → 用 \mathrm{} 写成正体
FUNCS_NAME = {"rank", "tr", "diag", "adj", "span", "sgn"}

# 字面 Unicode 上下标（agent 常直接打 ²、⁻、₁）
_SUPER = {"\u2070": "0", "\u00b9": "1", "\u00b2": "2", "\u00b3": "3",
          "\u2074": "4", "\u2075": "5", "\u2076": "6", "\u2077": "7",
          "\u2078": "8", "\u2079": "9", "\u207a": "+", "\u207b": "-",
          "\u207f": "n", "\u2071": "i"}
_SUB = {"\u2080": "0", "\u2081": "1", "\u2082": "2", "\u2083": "3",
        "\u2084": "4", "\u2085": "5", "\u2086": "6", "\u2087": "7",
        "\u2088": "8", "\u2089": "9", "\u208a": "+", "\u208b": "-"}
# 连续上下标必须整体成组。`⁻` 与 `⁷` 是**同一个指数**的两半（负号 + 数字），
# 逐字符各发一个 `^{-}` / `^{7}` 会得到 `10^{-}^{7}` —— 空上标组后紧跟第二个
# 上标，是非法 LaTeX，OMML 要么丢弃一半要么把 `7` 排成正文。所以先把连在一起的
# 上下标各自收成一个 run 再翻译。字符类里没有正则元字符，无需转义。
_SUPER_RUN_RE = re.compile("[" + "".join(_SUPER) + "]+")
_SUB_RUN_RE = re.compile("[" + "".join(_SUB) + "]+")

MATH_EXTRA = (set("+-−–×÷·⋅/|()[]{}^_ .,*'!;")
              | OP_CHARS | STRONG
              | {"\\", "&"}                       # LaTeX 环境写法：\begin{...}、& 对齐符
              | set(_SUPER) | set(_SUB)
              | {"\u00b0", "\u00b7", "\u22c5", "\u2032", "\u2033", "\u2034",
                 "\u00bd", "\u00bc", "\u00be", "\u2153", "\u2154"})

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\u3040-\u30ff\uac00-\ud7af]")
CJK_PUNCT = set("，。；：、！？（）【】「」『』《》〈〉…—～")

UNIT_LIST = sorted([
    "J/(mol·K)", "J/(kg·K)", "W/(m·K)", "J/mol",
    "kVA", "kvar", "kW", "kWh", "k\u03a9", "M\u03a9", "m\u03a9",
    "km/s", "m/s", "cm/s", "mm/s",
    "hPa", "kPa", "MPa", "GPa", "Pa",          # 压力：`p=101kPa` 这类不带斜杠的写法
    "\u03bcF", "pF", "nF", "mH", "\u03bcH", "kHz", "MHz", "Hz",
    "rad/s", "VA", "var", "mA", "\u03bcA", "kV", "mV", "ms", "\u03bcs", "ns",
    "\u03a9", "F", "H", "W", "V", "A", "s", "J", "T", "Wb",
], key=len, reverse=True)

_UNIT_ALT = "|".join(re.escape(u) for u in UNIT_LIST)
# 单位前必须是数字/右括号/右方括号（**不能是 }** —— 那是脚本结尾，A^{T}A 里的 A 是矩阵名）
# 后不能紧跟 ^ _ 字母数字
UNIT_RE = re.compile(
    rf"(?<=[0-9\)\]])\s*((?:{_UNIT_ALT})(?:\s*[·⋅]\s*(?:{_UNIT_ALT}))*)"
    rf"(?![_^{{A-Za-z0-9])"
)

# 幂次后面的单位（10^{-6}F、3.0×10^{8} m/s）：前面是脚本 }} 但内容是数字，仍算单位
#
# ★ 两种拼写必须给出同一个结果，否则 `3×10⁸m/s` 与 `3×10^{8}m/s` 会分叉：
#   * 花括号形态 `^{8}` —— 文档里本来就写成 LaTeX 的；
#   * **字面 Unicode 上标** `⁸` / `⁻⁷` —— agent 直接从物理题里拷出来的常见形态。
#   `_unitify` 跑在 `_escape_unicode` **之前**，所以此刻源串里还是字面形态；
#   只认 `^{n}` 会让字面写法漏掉这一步，单位被当变量排出斜体（`10^{8}s` 而非
#   `10^{8}\,\text{s}`）。这正是 R4 §3.2「幂次后单位」修复只修一半的残留。
#   字面字符类只收**数字**（⁰¹²³⁴⁵⁶⁷⁸⁹）：`⁺`/`⁻` 本身不是数值锚点。
#   ★ 从 `_SUPER` **派生**而不是手抄（REVIEW-CODE-R5 F-2）：下面 `_ANCHOR_RE`
#   用的是同一份 `_SUPER_DIGITS`。手抄两份的话，将来 `_SUPER` 增删数字键时
#   幂次锚点与单位锚点会静默分叉 —— 两个判据给出不同结果，正是本轮 parity
#   修复要消灭的那类缺陷。`_SUPER` 定义在本文件更上方（MATH_EXTRA 之前）。
_SUPER_DIGITS = "".join(c for c in _SUPER if _SUPER[c].isdigit())
POW_UNIT_RE = re.compile(
    rf"(\^\{{[-\u2212+]?\d+\}}|[{re.escape(_SUPER_DIGITS)}]+)"
    rf"\s*((?:{_UNIT_ALT})(?:\s*[·⋅]\s*(?:{_UNIT_ALT}))*)"
    rf"(?![_^{{A-Za-z0-9])"
)

# ---------------------------------------------------------------- 单位词表
# `/` 在量纲串里读作「每」（kg/m³ = 千克每立方米），只有在「量符号」之间才是分数线。
# 光判「两边是不是单位符号」是不够的 —— p=F/A、ρ=m/V、s/t 里两边也全是合法单位符号，
# 却是真分式。所以判定必须再带一层上下文锚点，规则见 _dimensionify()：
#
#   量纲串 := 数字? 单位组 ( '/' 单位组 )*         ← 至少一个 '/'
#   单位组 := 单位因子 ( ('·'|'⋅'|'*') 单位因子 )* | '(' 单位组 ')'
#   单位因子 := 单位符号 ('^' 整数 | 字面上标)?
#   单位符号 := UNITS 成员 | SI 词头 + UNITS 成员
#
# 成立条件（两条锚点任取其一）：
#   A. 数字锚点：串紧挨着一个数字/数字幂（`1.2kg/m³`、`3.0×10^{8} m/s²`）→ 测量值读法；
#   B. 显式登记：整条规范化后命中 COMPOUND_UNITS，并且左右是干净边界
#      （`mol/L`、`r/min` 单独出现可以；`E = V/m` 不行 —— 前面是 '='，读作真分式）。
# 都不满足时一律当分式，`v=s/t`、`p=F/A`、`ρ=m/V` 因此毫发无伤。

_UNIT_SEP = "\u00b7\u22c5\u2219*"          # 单位连乘号：· ⋅ ∙ *
# Inside a parenthesised unit group the ASCII full stop is also a product sign:
# `J/(mol.K)` is how a Chinese engineering document types `J/(mol·K)`.  It is
# deliberately *not* added to `_UNIT_SEP`: that set feeds `_CHAIN_MARKERS`, and
# a bare `.` there would turn `kg.m` into a dimension string on its own.  Here
# a full stop can only separate two unit factors inside `( )`, where a decimal
# point cannot appear.
_GROUP_SEP = _UNIT_SEP + "."
_WS_CHARS = " \t\u00a0"

SI_PREFIXES = frozenset([
    "da", "Y", "Z", "E", "P", "T", "G", "M", "k", "h", "d",
    "c", "m", "u", "\u03bc", "\u00b5", "n", "p", "f", "a", "z", "y",
])

# 单位词表：能独立作为一个「单位因子」出现的符号（含常用词头形式）。
# ★ 刻意不收的单字母：t（吨 / 时间符号）、d（天 / 微分符号）、c、a、e、M、U、Q、R、X。
#   它们在物理式子里当量符号的概率远高于当单位 —— `v=s/t` 必须保持真分式，就靠不收 t。
UNITS = frozenset([
    # 基本单位
    "m", "s", "A", "K", "mol", "cd", "kg",
    # 质量
    "g", "mg", "\u03bcg", "\u00b5g", "ug", "ng", "pg",
    # 长度
    "cm", "mm", "dm", "km", "\u03bcm", "\u00b5m", "um", "nm", "pm", "fm",
    # 时间（t / d 见上，刻意不收）
    "ms", "\u03bcs", "\u00b5s", "us", "ns", "ps", "fs", "min", "h",
    # 物质的量 / 体积
    "kmol", "mmol", "\u03bcmol", "\u00b5mol", "umol",
    "L", "l", "mL", "ml", "\u03bcL", "\u00b5L", "uL", "dL", "cL", "hL",
    # 力 / 压力
    "N", "kN", "mN", "\u03bcN", "\u00b5N",
    "Pa", "hPa", "kPa", "MPa", "GPa", "bar", "mbar", "atm", "mmHg",
    # 能量 / 功率 / 热量
    "J", "kJ", "MJ", "GJ", "mJ",
    "W", "kW", "MW", "GW", "mW", "\u03bcW", "\u00b5W",
    "Wh", "kWh", "MWh", "eV", "keV", "MeV", "GeV", "cal", "kcal",
    # 电 / 磁
    "V", "kV", "mV", "\u03bcV", "\u00b5V", "nV",
    "A", "mA", "\u03bcA", "\u00b5A", "uA", "kA", "nA",
    "F", "\u03bcF", "\u00b5F", "uF", "nF", "pF",
    "H", "mH", "\u03bcH", "\u00b5H", "uH", "nH",
    "C", "mC", "\u03bcC", "\u00b5C", "Ah", "mAh",
    "T", "mT", "\u03bcT", "nT", "Wb", "S", "Gs", "dB",
    "Hz", "kHz", "MHz", "GHz", "THz",
    "\u03a9", "\u2126", "k\u03a9", "M\u03a9", "m\u03a9",
    "VA", "kVA", "MVA", "var", "kvar",
    # 光 / 辐射 / 角 / 其他
    "lm", "lx", "Bq", "Gy", "Sv", "kat",
    "rad", "sr", "r", "rpm", "kn", "\u00c5", "Da",
])

# 显式登记的量纲串（规范化后比较：去空白、字面上标 → ^n、⋅/∗ → ·）。
# 只对「没有数字锚点」的场合生效，让正文里单摆着的 `mol/L`、`r/min` 也能认出来。
COMPOUND_UNITS = frozenset([
    "m/s", "m/s^2", "m/s^3", "km/s", "cm/s", "mm/s", "km/h", "m/h", "m/min",
    "rad/s", "rad/min", "r/min", "r/s", "rpm",
    "mol/L", "mmol/L", "umol/L", "mol/m^3", "mol/s", "mol/min",
    "g/mol", "kg/mol", "g/mL", "g/L", "kg/L", "mg/L", "mg/m^3", "g/cm^3",
    "kg/m^3", "kg/dm^3", "kg/L",
    "J/mol", "kJ/mol", "MJ/mol", "J/(mol\u00b7K)", "J/(kg\u00b7K)",
    "kJ/(kg\u00b7K)", "kJ/(mol\u00b7K)", "J/(m\u00b7K)",
    "W/m^2", "W/(m\u00b7K)", "W/(m^2\u00b7K)", "kW/m^2", "mW/cm^2",
    "A/m^2", "mA/cm^2", "A/cm^2", "A/mm^2",
    "A\u00b7h", "Ah", "mA\u00b7h", "mA\u00b7h/g", "kW\u00b7h", "kWh",
    "W\u00b7h", "V\u00b7A", "kVA",
    "N/m", "N\u00b7m", "N/m^2", "kN/m", "kN\u00b7m", "N/(m\u00b7s)",
    "V/m", "kV/m", "V/cm", "V/mm",
    "C/kg", "C/m^2", "C/m^3", "C/(kg\u00b7K)",
    "\u03a9\u00b7m", "\u03a9\u00b7cm", "\u03a9/cm", "S/m", "S/cm",
    "Pa\u00b7s", "mPa\u00b7s", "Pa/m",
    "L/min", "L/s", "L/h", "mL/min", "mL/s",
    "mol/(L\u00b7s)", "mol/(m^3\u00b7s)",
    "kg\u00b7m/s^2", "kg\u00b7m^2/s^2", "kg\u00b7m^2/s^3",
    "kg\u00b7m^2/(s^3\u00b7A)",
    "kcal/mol", "kcal/(kg\u00b7K)", "cal/(g\u00b7K)",
    "m^2/s", "m^2/s^2", "m^3/s", "m^3/min", "cm^3/s",
    "Wb/m^2", "kJ/mol", "mol/L",
])

# 单位符号 → LaTeX 命令（只处理会与 LaTeX/分词打架的两个希腊字形）
_UNIT_TEX = {
    "\u03a9": r"\Omega", "\u2126": r"\Omega",
    "\u03bc": r"\mu", "\u00b5": r"\mu",
}


def _sym_tex(sym: str) -> str:
    """单位符号 → LaTeX；命令后紧跟 ASCII 字母时补一个空格，免得粘成 \\mum。"""
    out = []
    for i, ch in enumerate(sym):
        cmd = _UNIT_TEX.get(ch)
        if cmd is None:
            out.append(ch)
            continue
        nxt = sym[i + 1:i + 2]
        out.append(cmd + (" " if nxt.isascii() and nxt.isalpha() else ""))
    return "".join(out)


# 关系符两边都要有「非运算符」的东西（挡掉 b=、=R、Uoc →，但允许 =∫、=Δ）
_BOTH_SIDES = re.compile(
    r"[^\s=≈≤≥≠∝±<>]\s*[=≈≤≥≠∝±<>]\s*[−+\-]?\s*[^\s=≈≤≥≠∝±<>]")

# 作者**显式**写的角标组：`S_{甲}` / `x^{乙}` / `m^{2}`。
# 这是「手动编辑等价物」——作者标注了才转，一个字符都不猜，所以误报恒为 0。
# 只认**带花括号**的写法（LaTeX 的标准形态，能容纳任意字符，含空格与标点）；
# 不带括号的 `x^2` 不在此列：它在正文里出现时意图不够明确。
_EXPLICIT_SCRIPT_RE = re.compile(r"[\^_]\{")

# 片段两端要裁掉的悬空运算符（箭头也算：手册里 "… → 下一步" 很常见）
_TRAIL_OP = "=+−-×÷·/,;<>≈≤≥≠∝±∓→←⇒"


# ---------------------------------------------------------------- 原子抽取
def _run_script(run) -> str | None:
    from docx.oxml.ns import qn
    rpr = run.find(qn("w:rPr"))
    if rpr is None:
        return None
    va = rpr.find(qn("w:vertAlign"))
    if va is None:
        return None
    val = va.get(qn("w:val"))
    return val if val in ("superscript", "subscript") else None


def paragraph_atoms(p_el) -> list[dict] | None:
    """把 w:p 展开成原子序列；遇到非纯 run 结构（超链接等）返回 None。"""
    from docx.oxml.ns import qn
    atoms: list[dict] = []
    for child in list(p_el):
        tag = child.tag
        if tag == qn("w:pPr"):
            continue
        if tag != qn("w:r"):
            return None                      # w:hyperlink / w:ins 等，跳过
        script = _run_script(child)
        rpr = child.find(qn("w:rPr"))
        # 按 run 内子元素的文档顺序展开（w:tab 与 w:t 的先后必须保序，
        # 否则 `…r²\t(2)` 会变成 `…r²(2)\t`，编号被吸进公式）
        for el in child:
            el_tag = el.tag
            if el_tag == qn("w:t"):
                for ch in (el.text or ""):
                    atoms.append({"ch": ch, "rpr": rpr, "script": script})
            elif el_tag == qn("w:tab"):
                atoms.append({"ch": "\t", "rpr": rpr, "script": None})
    return atoms


# ---------------------------------------------------------------- 片段判定
def is_math_char(ch: str) -> bool:
    if not ch:                       # special atom（\eqref 等）不占数学字符位
        return False
    # 先判数学集：U+00B7(·) 在本手册里是乘号，而它恰好也在中文标点表里
    if ch in MATH_EXTRA or ch in GREEK or ch in SYMBOL or ch in COMBINING:
        return True
    if ch in CJK_PUNCT or CJK_RE.match(ch):
        return False
    if ch.isascii() and ch.isalnum():
        return True
    cp = ord(ch)
    return (0x2190 <= cp <= 0x22ff) or (0x2a00 <= cp <= 0x2aff) or (0x1d400 <= cp <= 0x1d7ff)


# 显式角标组的内容上限与硬停字符。`S_{甲}` 里 `甲` 不是数学字符，但它属于
# 角标内容（任意字符都合法）；上限与句末标点只用于给**未闭合**的括号兜底，
# 防止一个手滑的 `_{` 把整段吞成公式。
# ⚠️ 口径（REVIEW-CODE-R5 K-2，别再写错）：判定是 `inside <= _EXPLICIT_SCRIPT_MAX`
# 且 inside 从 0 起数，所以组内实际**最多放行 25 个字符**（第 26 字起截断）。
# 超过上限被截断的 span 因 A-1 的平衡检查整段退回纯文本 —— 不丢字、不误排，
# 作者看到角标没渲染会发现内容太长或漏了 `}`。
_EXPLICIT_SCRIPT_MAX = 24
_SCRIPT_GROUP_STOP = set("。！？；\n\r\t")


def _open_script_group_len(sofar: str) -> int:
    r"""``sofar`` 停在未闭合的显式角标组里时，返回**组内已积累的字符数**。

    只认紧跟在 ``^`` / ``_`` 之后的 ``{``（``S_{甲}``）；其余花括号 ——
    ``\text{…}``、``\begin{…}``、普通分组 —— 各自由别处保护，不计入。
    返回 ``-1`` 表示当前不在任何显式角标组内。
    """
    stack: list[tuple[bool, int]] = []          # (是否由 ^/_ 开启, 内容起点)
    for i, ch in enumerate(sofar):
        if ch == "{":
            stack.append((i > 0 and sofar[i - 1] in "^_", i + 1))
        elif ch == "}" and stack:
            stack.pop()
    for opened_by_script, start in reversed(stack):
        if opened_by_script:
            return len(sofar) - start
    return -1


def _looks_like_date(s: str) -> bool:
    return bool(re.fullmatch(r"[\d\-./\s]+", s))


def find_math_spans(atoms: list[dict]) -> list[tuple[int, int]]:
    """返回 [(start, end)]（左闭右开，基于 atoms 下标）。"""
    spans: list[tuple[int, int]] = []
    i, n = 0, len(atoms)
    while i < n:
        if atoms[i].get("special"):      # \eqref / \label 等：不参与公式片段
            i += 1
            continue
        if not is_math_char(atoms[i]["ch"]):
            i += 1
            continue
        j = i
        while j < n:
            if atoms[j].get("special"):
                break
            if is_math_char(atoms[j]["ch"]) or atoms[j]["script"]:
                j += 1
                continue
            # 环境尚未闭合时，CJK 等字符不截断：\begin{cases} … 当 x>0 … \end{cases}
            sofar = "".join(a["ch"] for a in atoms[i:j])
            if sofar.count("\\begin{") > sofar.count("\\end{"):
                j += 1
                continue
            # 作者显式写的角标组未闭合时同样不截断：`S_{甲}` / `x^{乙}` 的
            # 角标内容本就是任意字符（含汉字），不该在这里被切成正文。
            # 上限与句末标点是给**手滑漏写 `}`** 的兜底，不是给正常用法的限制。
            inside = _open_script_group_len(sofar)
            if 0 <= inside <= _EXPLICIT_SCRIPT_MAX and atoms[j]["ch"] not in _SCRIPT_GROUP_STOP:
                j += 1
                continue
            break
        s, e = i, j
        # 裁掉两端空白与悬空运算符
        while s < e and (atoms[s]["ch"] == " " or atoms[s]["ch"] in _TRAIL_OP):
            s += 1
        while e > s and (atoms[e - 1]["ch"] == " " or atoms[e - 1]["ch"] in _TRAIL_OP):
            e -= 1
        text = "".join(a["ch"] for a in atoms[s:e])
        has_op = any(c in OP_CHARS for c in text)
        has_strong = any(c in STRONG for c in text)
        alnum = sum(1 for c in text if c.isalnum() or c in GREEK)
        # 显式角标标注（`S_{甲}` / `m^{2}`）本身就是「这是公式」的充分信号：
        # 它是作者的主动声明，不是我们的猜测 —— 所以它不必再靠关系符论证。
        # 这是「手动编辑等价物」路线能落地的一环，也是自动猜路线永远做不到的。
        explicit_script = bool(_EXPLICIT_SCRIPT_RE.search(text))
        # ★ span 文本自身不得停在未闭合的显式角标组内（REVIEW-CODE-R5 A-1）。
        #   未闭合 = 漏写 `}` 的笔误，或角标内容超过豁免上限被截断。这种 span
        #   若放行，下游 latex2mathml 会**自动补闭括号**：`_{` 两字符被吃进公式
        #   结构（document.xml 里消失）、其后 ≤25 字正文被排成巨型下角标 ——
        #   丢字 + 语义错排，比基线（整段纯文本）还差。所以整段退回纯文本，
        #   作者看到角标没渲染自然会发现漏了 `}`。
        #   合法的 `S_{甲}` 组已闭合，`_open_script_group_len` 返回 -1，不受影响。
        script_group_balanced = _open_script_group_len(text) == -1
        ok = (
            len(text) >= 2
            and not _looks_like_date(text)
            and alnum >= 1
            and (has_op or has_strong or explicit_script)
            and (bool(_BOTH_SIDES.search(text)) or explicit_script)
            and script_group_balanced
            # 纯数字 + 单位（如 6V、2Ω）不转
            and not re.fullmatch(r"[\d.]+\s*[A-Za-zΩμ]+", text)
        )
        if ok:
            spans.append((s, e))
        i = max(j, i + 1)
    return spans


# ---------------------------------------------------------------- 原子 → 线性串
def atoms_to_linear(atoms: list[dict], start: int, end: int) -> str:
    out: list[str] = []
    i = start
    while i < end:
        a = atoms[i]
        script = a["script"]
        if script:
            buf = []
            while i < end and atoms[i]["script"] == script:
                buf.append(atoms[i]["ch"])
                i += 1
            piece = "".join(buf)
            marker = "^" if script == "superscript" else "_"
            k = len(out) - 1
            while k >= 0 and out[k] == " ":
                k -= 1
            ins = "{" + piece + "}"
            if k >= 0:
                out.insert(k + 1, marker + ins)
            else:
                out.append(marker + ins)
            continue
        out.append(a["ch"])
        i += 1
    return "".join(out).strip()


# ---------------------------------------------------------------- 线性串 → LaTeX
_LETTER_CMD = re.compile(r"^\\[A-Za-z]+$")

# 裸 `^` / `_` 后面该吞进多少：**可选正负号 + 连续数字（含小数）**，或**单个字母**。
#
# 旧实现是 `\^([^\s{])` —— 只吃 1 个字符。于是 `10^-7` 被切成 `10^{-}7`：
# `-` 进了上标、`7` 留在正文，两者基线不同，Word 里看着像「中间多了个空格」。
# 这是主人报的原始 bug（HANDOFF-NEXT §2.1 根因 A）。
#
# 关键约束：`a^2b` 必须仍是 `a^{2}b`，不能变成 `a^{2b}`。所以这里**只吃
# 纯数字 token 或纯单字母**，绝不跨到下一个字母：
#   `10^-7` → 数字分支吃 `-7`      → `10^{-7}`
#   `10^+3` → 数字分支吃 `+3`      → `10^{+3}`
#   `a^2b`  → 数字分支吃 `2`（(?:\.\d+)? 不匹配 b） → `a^{2}b`
#   `x^n`   → 字母分支吃 `n`       → `x^{n}`
#   `x^{-7}`→ 已是 `{`，两个分支都不匹配 → 原样保留，不二次包裹
# 减号同时接受 U+2212（真减号），源文档两种都有。
SCRIPT_TOKEN = r"[-+\u2212]?\d+(?:\.\d+)?|[-+\u2212]?[A-Za-z]"
# 命令后的「可读空格」哨兵：分词时忽略（这样 \partial z 仍算邻接因子），
# 最终输出时再换成真正的空格（否则 \partialx 会被吃成一个命令）。
SP = "\x01"


def _escape_unicode(s: str) -> str:
    for comb, cmd in COMBINING.items():
        s = re.sub(r"(\\[A-Za-z]+|[A-Za-z0-9\)\]}])" + re.escape(comb),
                   lambda m, c=cmd: "\\" + c + "{" + m.group(1) + "}", s)
    # 连续上下标先合并成 run（`⁻⁷` → `^{-7}`），再逐字符处理余下的单个字符。
    # 不合并的话 `10⁻⁷` 会拆成 `10^{-}^{7}`（非法 LaTeX），见 _SUPER_RUN_RE 处的说明。
    s = _SUPER_RUN_RE.sub(
        lambda m: "^{" + "".join(_SUPER[c] for c in m.group(0)) + "}", s)
    s = _SUB_RUN_RE.sub(
        lambda m: "_{" + "".join(_SUB[c] for c in m.group(0)) + "}", s)
    out = []
    for ch in s:
        if ch in _SUPER:
            out.append("^{" + _SUPER[ch] + "}")     # 字面上标 ² → ^{2}
        elif ch in _SUB:
            out.append("_{" + _SUB[ch] + "}")       # 字面下标 ₁ → _{1}
        elif ch in GREEK:
            out.append(GREEK[ch] + SP)
        elif ch in SYMBOL:
            rep = SYMBOL[ch]
            out.append(rep + (SP if _LETTER_CMD.match(rep) else ""))
        elif ch in COMBINING:
            continue                      # 已处理
        else:
            out.append(ch)
    s = "".join(out)
    # 命令与上下标之间不能有空格：\omega_{0L}
    s = re.sub(r"(\\[A-Za-z]+)" + SP + r"(?=[_^])", r"\1", s)
    return re.sub(re.escape(SP) + r"{2,}", SP, s)


def _funcify(s: str) -> str:
    alt = "|".join(sorted(FUNCS, key=len, reverse=True))
    # ① 函数名后直接跟变量字母：必须补空格，否则 \sinx 是未定义命令
    s = re.sub(rf"(?<![\\A-Za-z])({alt})(?=[A-Za-z])",
               lambda m: "\\" + m.group(1) + " ", s)
    # ② 其余情况（后面是括号/上下标/空格/数字）
    s = re.sub(rf"(?<![\\A-Za-z])({alt})(?=[_^({{\[\s\\|]|\d)",
               lambda m: "\\" + m.group(1), s)
    # ③ LaTeX 没有 \tr/\rank 之类，写成正体
    alt_name = "|".join(sorted(FUNCS_NAME, key=len, reverse=True))
    s = re.sub(rf"(?<![\\A-Za-z])({alt_name})(?=[A-Za-z])",
               lambda m: "\\mathrm{" + m.group(1) + "} ", s)
    s = re.sub(rf"(?<![\\A-Za-z])({alt_name})(?=[_^({{\[\s\\|]|\d)",
               lambda m: "\\mathrm{" + m.group(1) + "}", s)
    return s


def _match_group(s: str, i: int, open_ch: str = "{", close_ch: str = "}") -> int:
    """s[i] == open_ch，返回配对 close_ch 之后的下标。"""
    depth = 0
    while i < len(s):
        if s[i] == open_ch:
            depth += 1
        elif s[i] == close_ch:
            depth -= 1
            if depth == 0:
                return i + 1
        i += 1
    return len(s)


def _protect_cmd(s: str, name: str, protect) -> str:
    r"""把 `\name{...}`（花括号可嵌套）整段交给 protect 换成占位符。

    与写死的 `\\text\{[^{}]*\}` 不同，这里按花括号配对扫描，所以
    `\mathrm{kg/m^{3}}` 这种带内层幂次的整段也能被原样托管。
    """
    out: list[str] = []
    i, n = 0, len(s)
    need = "\\" + name
    while i < n:
        if s[i] == "\\" and s.startswith(need, i):
            end_cmd = i + len(need)
            if end_cmd < n and s[end_cmd] == "{":
                end = _match_group(s, end_cmd)
                out.append(protect(s[i:end]))
                i = end
                continue
        out.append(s[i])
        i += 1
    return "".join(out)


def _sqrtify(s: str) -> str:
    """`√(...)` → `\\sqrt{...}`（括号可嵌套：√(1 + (y′)²)）。"""
    out: list[str] = []
    i = 0
    while True:
        j = s.find("\\surd", i)
        if j < 0:
            out.append(s[i:])
            break
        out.append(s[i:j])
        k = j + len("\\surd")
        while k < len(s) and s[k] in (" ", SP):
            k += 1
        if k < len(s) and s[k] == "(":
            end = _match_group(s, k, "(", ")")
            out.append("\\sqrt{" + s[k + 1:end - 1] + "}")
            i = end
        elif k < len(s) and s[k] == "{":
            end = _match_group(s, k, "{", "}")
            out.append("\\sqrt" + s[k:end])
            i = end
        elif k < len(s) and s[k].isalnum():
            out.append("\\sqrt{" + s[k] + "}")
            i = k + 1
        else:
            out.append("\\surd")
            i = k
    return "".join(out)


_CMD_RE = re.compile(r"\\[A-Za-z]+")
_NUM_RE = re.compile(r"\d+(?:\.\d+)?[A-Za-z]?")
# 整串就是一个化学式：元素符号（首字母大写）后跟可选数字，至少两段，且不含
# 运算符/括号/反斜杠。`H2O`、`CO2`、`H2SO4`、`C6H12O6`、`Na2CO3` 命中；
# `v0`、`R1=10`、`m1v1+m2v2` 不命中（含运算符或全小写）。
_CHEM_FORMULA_RE = re.compile(
    r"^(?:[A-Z][a-z]?\d*){2,}$")

# 乘号也是运算符：`1/2×1/2` 的分母必须止于 `2`，否则整串被吞进 \frac 的分母
# （旧实现把 `×` 当普通 atom，得到 \frac{\frac{1}{2\times 1}}{2}）。
# `×`/`·`/`⋅` 在 _fracify 之前已被 _escape_unicode 改写成命令（\times 等），
# 所以这两个命令名也必须在分词时被当成 op，否则修不到实际代码路径。
_OP_SPLIT = set("+=<>±∓,;|/÷×·⋅*")
_OP_CMD_NAMES = {"\\times", "\\cdot", "\\ast", "\\div"}   # 乘除法的命令形态


def _split_atoms(s: str) -> list[tuple[str, str]]:
    """把 LaTeX 串切成 (kind, text)，kind ∈ {atom, op, space}。"""
    out: list[tuple[str, str]] = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch in ("\x00", EPS):          # 占位符（\text{} / 环境）：整体一个 atom
            j = s.find(ch, i + 1)
            j = j + 1 if j > i else i + 1
            out.append(("atom", s[i:j]))
            i = j
            continue
        if ch == SP:                     # 命令后的可读空格哨兵：并入前一个 atom，
            if out and out[-1][0] == "atom":   # 不打断「邻接因子」判定
                out[-1] = ("atom", out[-1][1] + SP)
            else:
                out.append(("atom", SP))
            i += 1
            continue
        if ch == "\\":
            m = _CMD_RE.match(s, i)
            if not m:
                out.append(("atom", ch))
                i += 1
                continue
            j = m.end()
            if m.group(0) in _OP_CMD_NAMES and j < n and s[j] != "{":
                # \times 这类命令是运算符，不能当成邻接因子（否则 / 的分母会
                # 越过它继续右吃，`1/2\times 1/2` 变成嵌套分式）
                out.append(("op", m.group(0)))
                i = j
                continue
            if j < n and s[j] == "{":
                k = _match_group(s, j)
                out.append(("atom", s[i:k]))
                i = k
            else:
                out.append(("atom", s[i:j]))
                i = j
        elif ch in "^_":
            if i + 1 < n and s[i + 1] == "{":
                k = _match_group(s, i + 1)
                piece = s[i:k]
                i = k
            else:
                piece = s[i:i + 2]
                i += 2
            if out and out[-1][0] == "atom":
                out[-1] = ("atom", out[-1][1] + piece)
            else:
                out.append(("atom", piece))
        elif ch == "{":
            k = _match_group(s, i)
            out.append(("atom", s[i:k]))
            i = k
        elif ch == "(":
            k = _match_group(s, i, "(", ")")
            out.append(("atom", s[i:k]))
            i = k
        elif ch == " ":
            out.append(("space", ch))
            i += 1
        elif ch in _OP_SPLIT:
            out.append(("op", ch))
            i += 1
        else:
            m = _NUM_RE.match(s, i)
            if m and (m.end() - i > 1 or ch.isdigit()):
                out.append(("atom", m.group(0)))
                i = m.end()
            else:
                out.append(("atom", ch))
                i += 1
    return out


def _strip_parens(x: str) -> str:
    """整个操作数就是一对括号时去掉它（进 \\frac 的 {} 后不需要）。"""
    if x.startswith("(") and x.endswith(")") and _match_group(x, 0, "(", ")") == len(x):
        return x[1:-1]
    return x


def _is_whole_group(text: str) -> bool:
    """text 本身是不是「一个完整的分组」（而不是 分组+上下标 之类）。"""
    if not text or text[:1] not in "({":
        return False
    close = ")" if text[0] == "(" else "}"
    return _match_group(text, 0, text[0], close) == len(text)


# 这些「裸」大运算符不算分式操作数的一部分：∫dx/(1+x²) 的分子是 dx，不是 ∫dx
_BARE_BIGOPS = {r"\int", r"\oint", r"\iint", r"\iiint", r"\sum", r"\prod",
                r"\lim", r"\limsup", r"\liminf"}


def _fracify(s: str) -> str:
    r"""`a/b` → `\frac{a}{b}`。

    操作数按「邻接因子」扩展：`Gm_{1}m_{2}/r^{2}` → `\frac{Gm_{1}m_{2}}{r^{2}}`、
    `dp/dt` → `\frac{dp}{dt}`、`∂z/∂x` → `\frac{\partial z}{\partial x}`、
    `m^{2}/kg^{2}` → `\frac{m^{2}}{kg^{2}}`。
    遇到运算符即停（空格不算分隔），所以 `a + b/c` 不会把 `a` 拉进分子。
    """
    toks = _split_atoms(s)

    # 先递归处理分组内部（只处理「整个 atom 就是一个完整分组」的情况）
    for idx, (kind, text) in enumerate(toks):
        if kind != "atom":
            continue
        if _is_whole_group(text):
            toks[idx] = ("atom", text[0] + _fracify(text[1:-1]) + text[-1])
            continue
        m = _CMD_RE.match(text) if text.startswith("\\") else None
        if m and m.end() < len(text):
            rest = text[m.end():]
            if _is_whole_group(rest):
                toks[idx] = ("atom", text[:m.end()] + "{" +
                             _fracify(rest[1:-1]) + "}")

    changed = True
    while changed:
        changed = False
        for i, (kind, text) in enumerate(toks):
            if kind != "op" or text != "/":
                continue
            p = i - 1
            while p >= 0 and toks[p][0] == "space":
                p -= 1
            q = i + 1
            while q < len(toks) and toks[q][0] == "space":
                q += 1
            if p < 0 or q >= len(toks):
                continue
            if toks[p][0] != "atom" or toks[q][0] != "atom":
                continue
            ll = p
            L = toks[ll][1]
            while ll - 1 >= 0 and toks[ll - 1][0] == "atom":
                prev = toks[ll - 1][1].replace(SP, "").strip()
                if prev in _BARE_BIGOPS and not _is_whole_group(toks[ll][1]):
                    # 裸 ∫/∑ 后面直接跟 dx 这类：不把它拉进分子
                    # （∫dx/(1+x²) → \int \frac{dx}{…}；但 Σ(E/R)/Σ(1/R) 会带上）
                    break
                ll -= 1
                L = toks[ll][1] + L
            rr = q
            R = toks[rr][1]
            while rr + 1 < len(toks) and toks[rr + 1][0] == "atom":
                rr += 1
                R += toks[rr][1]
            L, R = _strip_parens(L), _strip_parens(R)
            # 操作数两端可能挂命令后的空格哨兵（`\times` 之后就是这种），落进
            # \frac 的花括号会变成多余空格。只削两端：中间那个哨兵是命令与
            # 后续字符的分隔（`\omega L` 去掉就粘成未定义的 \omegaL）。
            L = L.strip(" " + SP)
            R = R.strip(" " + SP)
            toks[ll:rr + 1] = [("atom", r"\frac{" + L + "}{" + R + "}")]
            changed = True
            break
    return "".join(t for _, t in toks)


# ------------------------------------------------------- 量纲串判定（`/` = 每）
def _read_exponent(s: str, i: int) -> tuple[int, str] | None:
    """s[i] 起读幂次，返回 (结束下标, 指数文本)；不是幂次返回 None。

    三种写法都吃：LaTeX 式 `^{3}` / `^3`、字面 Unicode 上标 `³` / `⁻³`，
    以及**紧贴单位符号的裸整数** `kg/m3` / `m/s2`。

    最后一种不是规范写法，但它是人真正打字的样子——`³` 在多数输入法里
    打不出来，而 `kg/m3` 在工程文档、教辅、试卷里随处可见。它由调用方
    `_parse_unit_item` 控制是否开启：只有当这个整数**确实是一段量纲串的
    幂次**时才当幂次读，见那里的说明。
    """
    if i >= len(s):
        return None
    if s[i] == "^":
        j = i + 1
        if j < len(s) and s[j] == "{":
            k = _match_group(s, j)
            body = s[j + 1:k - 1].strip()
        else:
            m = re.match(r"[-\u2212+]?\d+", s[j:])
            if not m:
                return None
            body = m.group(0)
            k = j + m.end()
        body = body.replace("\u2212", "-")
        if not re.fullmatch(r"[-+]?\d+", body):
            return None
        return k, body
    j = i
    buf: list[str] = []
    while j < len(s) and s[j] in _SUPER:
        buf.append(_SUPER[s[j]])
        j += 1
    if not buf:
        return None
    return j, "".join(buf)


# A bare ASCII digit run right after a unit symbol, e.g. the ``3`` of ``kg/m3``.
# Bounded so that ``m3`` in ``m30`` (a typo or an unrelated token) still reads
# as an exponent, but a longer number does not get chopped arbitrarily.
_BARE_EXP_RE = re.compile(r"[-\u2212]?[1-9]\d?")


def _unit_symbol_at(s: str, i: int) -> int:
    """s[i:] 起的最长「单位符号」长度（UNITS 成员，或 SI 词头 + UNITS 成员）。"""
    best = 0
    for u in _UNITS_BY_LEN:
        if s.startswith(u, i):
            best = len(u)
            break
    for p in _PREFIXES_BY_LEN:
        if not s.startswith(p, i):
            continue
        j = i + len(p)
        if j == i:
            continue
        for u in _UNITS_BY_LEN:
            if s.startswith(u, j):
                best = max(best, j + len(u) - i)
                break
    return best


def _skip_ws(s: str, i: int) -> int:
    while i < len(s) and s[i] in _WS_CHARS:
        i += 1
    return i


def _parse_unit_item(s: str, i: int) -> tuple[int, str, bool] | None:
    """s[i:] 起的一个「单位因子」：单位符号 + 可选幂次。(结束, LaTeX, 有幂次)。"""
    n = _unit_symbol_at(s, i)
    if not n:
        return None
    sym = s[i:i + n]
    j = i + n
    e = _read_exponent(s, j)
    if e is None:
        # Bare ASCII exponent: ``kg/m3``, ``m/s2``. Accepted only when the
        # digits are not glued to a following alphanumeric, so ``m30x`` or
        # ``kg/m3s`` do not silently become ``m^{30}`` / ``m^{3}s``.
        m = _BARE_EXP_RE.match(s, j)
        if m and (m.end() >= len(s) or not s[m.end()].isalnum()):
            return m.end(), _sym_tex(sym) + "^{" + m.group(0).replace("\u2212", "-") + "}", True
        return j, _sym_tex(sym), False
    k, exp = e
    return k, _sym_tex(sym) + "^{" + exp + "}", True


def _parse_unit_group(s: str, i: int, sep: str = _UNIT_SEP) -> tuple[int, str, bool] | None:
    """s[i:] 起的一个「单位组」（可带括号、可含 · 连乘）。

    *sep* is the set of product signs accepted at this nesting level;
    ``_parse_unit_group`` passes ``_GROUP_SEP`` inside parentheses so that
    ``J/(mol.K)`` reads as a dimension string while a bare ``kg.m`` does not.
    """
    if i < len(s) and s[i] == "(":
        inner = _parse_unit_group(s, i + 1, _GROUP_SEP)
        if inner is None:
            return None
        j, tex, has_exp = inner
        k = _skip_ws(s, j)
        if k >= len(s) or s[k] != ")":
            return None
        end = k + 1
        e = _read_exponent(s, end)
        if e is not None:
            end, exp = e
            return end, "(" + tex + ")^{" + exp + "}", True
        return end, "(" + tex + ")", has_exp
    item = _parse_unit_item(s, i)
    if item is None:
        return None
    j, tex, has_exp = item
    parts = [tex]
    while True:
        k = _skip_ws(s, j)
        if k >= len(s) or s[k] not in sep:
            break
        nxt = _parse_unit_item(s, _skip_ws(s, k + 1))
        if nxt is None:
            break
        j, nxt_tex, nxt_exp = nxt
        sep_ch = s[k]
        # keep `.` as a literal dot, render the real product signs as \cdot
        parts.append(("." if sep_ch == "." else r"\cdot ") + nxt_tex)
        has_exp = has_exp or nxt_exp
    return j, "".join(parts), has_exp


# 数字锚点：串前面紧挨着一个数字 / 一个数字指数的右花括号 / 一个右括号 /
# 一段字面 Unicode 上标数字（`3×10⁸m/s` 里那个 `⁸`）。
# 最后一种和 `^{8}` 是同一件事的两种写法：它们都说明「左边这个数是数值的幂」，
# 因此右边的单位串读作测量值而不是分式。`_find_dimension_spans` 跑在
# `_escape_unicode` 之前，源文档里 `⁸` 还是字面形态，两种形态都要认——否则
# `x⁸m/s` 与 `x^{8}m/s` 会给出不同结果，而 ASCII/Unicode 两种拼写必须一致
# （见 tests/test_units_ascii.py 的同名约定）。
# 字符类只收数字：`⁺`/`⁻` 单独出现时不能当锚点（没有数值可锚），`ⁿ`/`ⁱ`
# 是字母上标，同样不是数值。`10⁻⁷` 里 `⁻` 后面还跟着 `⁷`，run 从 `⁷` 起算即可。
# 第 1 组捕获「数字与量纲串之间的空白」——这段空白要一起吃掉，
# 否则 `8.31 J/(mol·K)` 会留下一个旧实现没有的空格（既有用例锁的是无空格形态）。
# `_SUPER_DIGITS` 已在文件上方（`POW_UNIT_RE` 旁）从 `_SUPER` 派生，两处共用一份。
_ANCHOR_RE = re.compile(
    r"(?:\)|\^\{\s*[-\u2212+]?\d+\s*\}|[" + re.escape(_SUPER_DIGITS) + r"]+|\d)(\s*)$")
# 量纲串尾部不能紧跟这些字符（否则只是更长 token 的一段）
_TAIL_BLOCK_RE = re.compile(r"[_^{A-Za-z0-9]")
# 无数字锚点时，左边界不许是这些字符（`E = V/m` 的 `=`、`p=F/A` 的 `=` 都在内）
_LEFT_BLOCK = set("=+-*/^_<>'\u00b1\u2213\u00d7\u00f7\u00b7\u22c5\u2219,;|)}]")
# 一条链至少要有一个「分隔标记」：分数线或连乘号
_CHAIN_MARKERS = "/" + _UNIT_SEP


def _normalize_chain(text: str) -> str:
    """规范化用于与 COMPOUND_UNITS 比对：去空白、字面上标 → ^n、⋅/∗ → ·。"""
    out: list[str] = []
    for ch in text:
        if ch in _WS_CHARS:
            continue
        if ch in _SUPER:
            out.append("^" + _SUPER[ch])
        elif ch == "\u2212":
            out.append("-")
        elif ch in "\u22c5\u2219*":
            out.append("\u00b7")
        else:
            out.append(ch)
    return re.sub(r"\^\{([^}]*)\}", r"^\1", "".join(out))


def _left_boundary_ok(s: str, start: int) -> bool:
    """量纲串左边是不是干净边界（前一个非空白字符不是运算符/字母数字/右括号）。"""
    p = start - 1
    while p >= 0 and s[p] in _WS_CHARS:
        p -= 1
    if p < 0:
        return True
    ch = s[p]
    return not (ch.isalnum() or ch in _LEFT_BLOCK)


def _anchor_span_start(s: str, start: int, end: int) -> int | None:
    """量纲串的成立条件（见文件头规则说明）；成立时返回含左侧空白的起始下标。"""
    if end < len(s) and _TAIL_BLOCK_RE.match(s[end]):
        return None
    m = _ANCHOR_RE.search(s[:start])
    if m:
        return m.start(1)
    if not _left_boundary_ok(s, start):
        return None
    if _normalize_chain(s[start:end]) in COMPOUND_UNITS:
        return start
    return None


def _find_dimension_spans(s: str) -> list[tuple[int, int, str, str]]:
    """→ [(替换起点, 替换终点, 渲染好的 LaTeX, 原文片段)]，按出现顺序、互不重叠。

    替换起点可能落在量纲串左侧的空白上（数字与单位之间那个空格）。
    """
    spans: list[tuple[int, int, str, str]] = []
    i, n = 0, len(s)
    while i < n:
        got = _parse_unit_group(s, i)
        if got is None:
            i += 1
            continue
        end, tex, has_exp = got
        ok = True
        while True:
            k = _skip_ws(s, end)
            if k >= n or s[k] != "/":
                break
            nxt = _parse_unit_group(s, _skip_ws(s, k + 1))
            if nxt is None:
                ok = False
                break
            end, nxt_tex, nxt_exp = nxt
            tex = tex + "/" + nxt_tex
            has_exp = has_exp or nxt_exp
        raw = s[i:end]
        if not ok or not any(c in raw for c in _CHAIN_MARKERS):
            i += 1
            continue
        span_start = _anchor_span_start(s, i, end)
        if span_start is None:
            i += 1
            continue
        chain = _tidy_chain(raw)
        if has_exp:
            rendered = r"\,\mathrm{" + tex + "}"
        else:
            rendered = r"\,\text{" + chain + "}"
        spans.append((span_start, end, rendered, chain))
        i = end
    return spans


def _tidy_chain(text: str) -> str:
    """去掉 `/` 与连乘号两侧的空格，便于直接落进 \\text{}。"""
    text = re.sub(r"\s*/\s*", "/", text)
    return re.sub(r"\s*([\u00b7\u22c5\u2219])\s*", r"\1", text)


def is_dimension_string(text: str) -> bool:
    """整段 `text` 是不是量纲串（`/` 读作「每」）。供测试与报告判定。

    >>> is_dimension_string("kg/m\u00b3"), is_dimension_string("v=s/t")
    (True, False)
    """
    spans = _find_dimension_spans(text)
    return bool(spans) and spans[0][0] == 0 and spans[-1][1] == len(text)


def unit_slash_chains(linear: str) -> list[str]:
    """linear 串里被判为量纲串的原文片段（按出现顺序）。"""
    return [raw for _s, _e, _r, raw in _find_dimension_spans(linear)]


def _dimensionify(s: str) -> str:
    r"""把量纲串整段摘出来，挡住后面的分式改写。

    无幂次 → `\,\text{kg/m}`（与既有 UNIT_RE 的形态一致，`\text{}` 里塞不了花括号）；
    有幂次 → `\,\mathrm{m/s^{2}}`（幂次要落成真 OMML 上标）。
    """
    spans = _find_dimension_spans(s)
    if not spans:
        return s
    out: list[str] = []
    pos = 0
    for span_start, end, rendered, _raw in spans:
        out.append(s[pos:span_start])
        out.append(rendered)
        pos = end
    out.append(s[pos:])
    return "".join(out)


_UNITS_BY_LEN = sorted(UNITS, key=len, reverse=True)
_PREFIXES_BY_LEN = sorted(SI_PREFIXES, key=len, reverse=True)


def _unitify(s: str) -> str:
    # 0) 含 `/` 的量纲串先整体识别 —— 单位串里的 `/` 是「每」，不是分数线。
    #    必须先于分式改写，且必须先于下面两条白名单正则（否则 kg 会被切走）。
    s = _dimensionify(s)
    s = UNIT_RE.sub(lambda m: r"\,\text{" + m.group(1) + "}", s)
    # 幂次后的单位（10^{-6}F / 3.0×10^{8} m/s）——} 后面接数字脚本时仍按单位处理
    s = POW_UNIT_RE.sub(
        lambda m: m.group(1) + r"\,\text{" + m.group(2) + "}", s)
    return s


_ENV_RE = re.compile(r"\\begin\{(\w+)\}.*?\\end\{\1\}", re.S)
EPS = "\x02"          # 环境占位符的定界符（\x00 被 \text{} 占位用了）

# ------------------------------------------------------------ 手写矩阵文本
# agent 写矩阵的另一种现实形态：`[a, b; c, d]`（; 分行 , 分列）
_MATRIX_WRAP = {"[": ("bmatrix", "]"), "(": ("pmatrix", ")")}


def _matrix_body(body: str) -> list[list[str]] | None:
    """括号内文 → 行×列；不是规整矩阵就返回 None（保守回退原文）。"""
    if ";" not in body:
        return None
    # 嵌套括号 / 命令 / 环境会让 , ; 的语义歧义 → 一律不转
    if any(c in body for c in "([{\\"):
        return None
    rows: list[list[str]] = []
    for raw in body.split(";"):
        cells = [c.strip() for c in raw.split(",")]
        if any(not c for c in cells):
            return None
        rows.append(cells)
    if len(rows) < 2 or len({len(r) for r in rows}) != 1:
        return None
    return rows


def _matrixify(s: str) -> str:
    r"""`[a, b; c, d]` → `\begin{bmatrix}a & b \\ c & d\end{bmatrix}`。

    只转「含 ; 行分隔（≥2 行）且各行列数一致」的方括号/圆括号组；
    单行 `[a, b]`（区间、坐标、函数参数）不受影响。
    """
    out: list[str] = []
    i, n = 0, len(s)
    while i < n:
        ch = s[i]
        if ch in _MATRIX_WRAP:
            close = _MATRIX_WRAP[ch][1]
            end = _match_group(s, i, ch, close)
            if end < n or (end == n and s.endswith(close)):
                body = s[i + 1:end - 1]
                grid = _matrix_body(body)
                if grid is not None:
                    env = _MATRIX_WRAP[ch][0]
                    rows = [" & ".join(r) for r in grid]
                    out.append("\\begin{" + env + "}"
                               + " \\\\ ".join(rows) + "\\end{" + env + "}")
                    i = end
                    continue
        out.append(ch)
        i += 1
    return "".join(out)


def _subscriptify(s: str) -> str:
    r"""`v0` → `v_{0}`：变量名后面紧跟的**单个数字**通常是下标。

    物理卷里 `v₀`、`R₁`、`U₁`、`m₂` 天天出现，而源文档几乎只会写成
    `v0` / `R1` / `m2`（真下标 run 才走 `_run_script` 那条路）。不补这一步，
    `v0` 会原样输出成两个同高字符，看起来像 v 乘以 0，物理含义全错。

    判据刻意保守，因为**化学式与函数名都在同一条代码路径上**：

    * **小写**单字母 + 单个数字 → 转（`v0`、`m1v1`、`x0`、`a1`、`u2`）。
      物理/数学的变量几乎都是小写，这条覆盖绝大部分需求。
    * **大写**单字母 + 单个数字 → 只在数字后是运算符/边界时转（`R1=10`、
      `U1=`、`T1=`），从而**放过化学式**：`H2SO4` 的 `H2` 后面是 `S`、
      `CO2` 的 `O` 前是 `C`，都不动 —— 化学式在 spec 里本来就写成真下标
      （`H_{2}O`），不该被这里二次改写。
    * 字母前不能是字母/数字之外的命令成分：`\log2x`、`\lg2`、`\sin2x`
      经 `_funcify` 后已带反斜杠，前置断言把它们排除。
    * `e`/`E` 一律不动（`e2` 多为科学记数、`E1` 多为编号）。
    * **整串是化学式就不动**（`H2O`、`CO2`、`H2SO4`、`C6H12O6`）。化学式在 spec
      里本来就写成真下标（`H_{2}O`），走到这里的是漏写的；与其猜哪个数字是下标，
      不如整串放过，让人去补 —— 猜错（`C6H12O_{6}`）比不转更难发现。
    * **紧跟在字母后面的不转**（`IR1`、`I1R1`）：`IR₁` 与 `I·R1` 有歧义，
      要靠人写 `I_{1}R_{1}` 才明确。这条保守规则宁可不动，也不猜。

    幂次（`v0^2`）由后置的 `^` 规则处理，这里只碰底数。
    """
    if _CHEM_FORMULA_RE.match(s.replace(" ", "")):
        return s
    # 小写：前不接字母/反斜杠命令；数字后不再跟数字
    s = re.sub(r"(?<![\\A-Za-z])(?<![\\A-Za-z]\\{)([a-df-z])(\d)(?!\d)",
               r"\1_{\2}", s)
    # 大写：额外要求数字后是运算符或边界，放过 H2SO4 / CO2 这类化学式片段
    s = re.sub(r"(?<![\\A-Za-z])(?<![\\A-Za-z]\\{)([A-DF-Z])(\d)(?!\d)"
               r"(?=[\s=+\-−*/×·,;)}\]|]|$)",
               r"\1_{\2}", s)
    return s


# ---------------------------------------------------------------- 中文角标
# `S甲` / `P甲=P乙` / `v甲>v乙` 里的「甲乙」是下角标（主人 2026-11-06 反馈）。
# 但**绝不能默认开**：正文里 `m与`、`m高`、`t时`、`图甲`、`表乙` 长得一模一样，
# 自动猜会把正文打成下标。所以做成显式开关，默认 False —— 与 HANDOFF-NEXT §3.3
# 的首选方案（spec 里显式标注 vert_align）天然对齐：那条路产出真
# `w:vertAlign="subscript"` run；这条只服务「线性串里已经写成 `S甲`」的场景。
CN_SUBSCRIPT = False

# 汉字黑名单：紧跟在字母后面时**几乎必然是正文连接词/量词/方位词**，不是角标。
# 来源：HANDOFF-NEXT §3.3 给的那串 + 本轮实测补充的虚词与程度副词。
# 判据是「这个字接在变量后面读起来像不像一句中文」—— 像，就不转。
CN_SUBSCRIPT_BLACKLIST = frozenset(
    # —— HANDOFF §3.3 原表 ——
    "与和的时高长宽为是中内外上下后前处点段图表式题解答则若当已未"
    # —— 连接词 / 介词 / 助词 ——
    "比及或且即于由从向对把被让使令得地着过们这那其之所"
    # —— 连词 / 副词 / 语气 ——
    "以因如虽但却而又也都还就才只更最很太极再又并同"
    # —— 常见谓语 / 判断词（`S是`、`S有`、`S为`） ——
    "有在无不可会要应需能"
    # —— 常见量名词（`S值`、`S数`、`S量` 多半是正文） ——
    "值数量度数率差比积合"
)


def _cn_subscriptify(s: str) -> str:
    r"""`X汉` → `X_{汉}`：单个 ASCII 字母 + 单个汉字 → 中文下角标。

    前置断言 `(?<![A-Za-z0-9\\])` 保证「汉字前是一个**独立的** ASCII 字母」。
    四种必须放过的形态，都被断言挡住：

    * `ab甲` —— 字母是更长单词的一部分（前面还是字母）；
    * `28M战术` —— 字母是型号/单位代号的一部分：**多位数字**紧跟单个字母
      是 `28M`、`4G`、`100W` 这类编号，不是「28 乘 M」；
    * `\alpha甲` —— 命令尾字母。

    而 `2v乙` 要转：**单个数字**紧跟字母是系数（`2v`、`3x`），中文里
    「2v乙」读作 2 倍 v_乙。所以数字那一档只排除 2 位以上（`(?<!\d)` 之外
    再用 `(?<!\d\d)`，并把「单个数字 + 字母」显式放行）。

    汉字走 `CN_SUBSCRIPT_BLACKLIST` 之外才转。

    例：`S甲` → `S_{甲}`；`P甲=P乙` → `P_{甲}=P_{乙}`；`v甲=2v乙` →
        `v_{甲}=2v_{乙}`；`m与` / `m高` / `t时` / `图甲` / `28M战术` 不动。
    """
    out = []
    i = 0
    n = len(s)
    while i < n:
        ch = s[i]
        if (ch.isascii() and ch.isalpha() and i + 1 < n
                and "\u4e00" <= s[i + 1] <= "\u9fff"
                and s[i + 1] not in CN_SUBSCRIPT_BLACKLIST):
            prev = s[i - 1] if i else ""
            # Reject `ab甲` (word tail) and `\alpha甲` (command tail).
            if prev and (prev.isascii() and (prev.isalpha() or prev == "\\")):
                out.append(ch)
                i += 1
                continue
            # Reject `28M战术` / `4G通信`: a digit run before a *capital* letter
            # is a model/unit designation (28M, 4G, 5K, 3D), never a
            # coefficient, because a coefficient in Chinese physics prose is
            # lowercase (`2v乙`, `3x甲`).  Two or more digits block either case.
            if prev.isdigit():
                j = i - 1
                ndig = 0
                while j >= 0 and s[j].isdigit():
                    ndig += 1
                    j -= 1
                if ndig >= 2 or ch.isupper():
                    out.append(ch)
                    i += 1
                    continue
            out.append(ch + "_{" + s[i + 1] + "}")
            i += 2
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def linear_to_latex(s: str, cn_subscript: bool | None = None) -> str:
    # 0) latex2mathml 对 align/gather 会**自动插入行编号** (1)(2)（LaTeX 语义），
    #    没有 \tag 的多行环境一律换成星号版本（不编号）；aligned 同样映射到 align*。
    #    有 \tag 的由 docx_mathify 走「对齐编号表格」，不经过这里。
    s = s.replace("\\begin{aligned}", "\\begin{align*}").replace(
        "\\end{aligned}", "\\end{align*}")
    if "\\tag" not in s:
        s = s.replace("\\begin{align}", "\\begin{align*}").replace(
            "\\end{align}", "\\end{align*}")
        s = s.replace("\\begin{gather}", "\\begin{gather*}").replace(
            "\\end{gather}", "\\end{gather*}")
    # 0.5) 手写矩阵文本 [a, b; c, d] → bmatrix 环境（先于一切改写，原始 Unicode 上做）
    s = _matrixify(s)
    # 1) 先在原始 Unicode 上吃单位（30Ω / 1μF / 314 rad/s / 7.9 km/s）
    s = _unitify(s)
    # 2) 保护 \text{...}，避免单位里的字母被当成分式/命令
    protected: list[str] = []

    def _protect_text(txt: str) -> str:
        protected.append(txt)
        return "\x00" + str(len(protected) - 1) + "\x00"

    s = re.sub(r"\\text\{[^{}]*\}", lambda m: _protect_text(m.group(0)), s)
    # 2b) \mathrm{...}（量纲串的输出形态）整段当不透明块：别让 _funcify 把里面的
    #     min 改写成 \min，也别让 _fracify 伸进显式正体组。花括号按配对扫描。
    s = _protect_cmd(s, "mathrm", _protect_text)
    # 3) Unicode → 命令
    s = _escape_unicode(s)
    # 4) 保护 \begin{...}...\end{...}（矩阵/分段/多行环境：内容已是合法 LaTeX，
    #    不能再被 sqrt/func/frac 改写；& 和 \\ 交给 latex2mathml 自己解析）
    envs: list[str] = []

    def _protect_env(m):
        envs.append(m.group(0))
        return f"{EPS}{len(envs) - 1}{EPS}"

    s = _ENV_RE.sub(_protect_env, s)
    # 5) 根号、函数名、分式
    s = _sqrtify(s)
    s = _funcify(s)
    # 函数名此时已带反斜杠（\lg、\log），_subscriptify 的前置断言会把它们排除，
    # 所以 `v0`→`v_{0}` 生效而 `\lg2`/`\log2x` 不受影响。
    s = _subscriptify(s)
    s = _fracify(s)
    # 5b) 中文下角标（`S甲`）。**默认关**：显式开关或模块级 CN_SUBSCRIPT 才生效。
    #     放在占位符还原之前 —— `\text{…}`/`\mathrm{…}`/环境还都是 \x00N\x00 占位符，
    #     正文里的「字母+汉字」（`m与`、`图甲`）不会被打成下标。
    if CN_SUBSCRIPT if cn_subscript is None else cn_subscript:
        s = _cn_subscriptify(s)
    # 6) 还原占位符
    s = re.sub(r"\x00(\d+)\x00", lambda m: protected[int(m.group(1))], s)
    s = re.sub(r"\x02(\d+)\x02", lambda m: envs[int(m.group(1))], s)
    s = s.replace(SP, " ")                 # 命令后的可读空格（此时才落成真空格）
    s = re.sub(r" {2,}", " ", s).strip()
    s = re.sub(r"\^(" + SCRIPT_TOKEN + ")", r"^{\1}", s)
    s = re.sub(r"_(" + SCRIPT_TOKEN + ")", r"_{\1}", s)
    # 收尾：去掉 { } ( ) ^ _ , 之前的空格（LaTeX 不需要，且影响可读性）
    s = re.sub(r"\s+(?=[{}()^_,])", "", s)
    return s.strip()


def repair_merged_scripts(latex: str) -> tuple[str, list[str]]:
    """修正源文档里「把后面正常字符一起设成上下标」的排版错误。

    例：I 的指数被写成 "2R"（原文 I²R）→ `I^{2R}` 修成 `I^{2}R`；
        R 的下标被写成 "1R2"（原文 R₁R₂）→ 修成 `R_{1}R_{2}`。
    只在「数字 + 大写字母开头」这种明显不可能作为单一上下标的内容上动手。
    """
    fixes: list[str] = []

    def repl(m):
        marker, d1, letters, d2 = m.group(1), m.group(2), m.group(3), m.group(4)
        out = f"{marker}{{{d1}}}{letters}"
        if d2:
            out += f"_{{{d2}}}"
        fixes.append(m.group(0) + " -> " + out)
        return out

    pat = re.compile(r"([_^])\{(\d+)([A-Z][A-Za-z]*)(\d+)?\}")
    return pat.sub(repl, latex), fixes


# ---------------------------------------------------------------- 段落级
def split_paragraph(atoms: list[dict]) -> list[dict]:
    """→ [{'kind':'text'|'math', 'atoms':…, 'latex':…}]"""
    spans = find_math_spans(atoms)
    segs: list[dict] = []
    pos = 0
    for (s, e) in spans:
        if s > pos:
            segs.append({"kind": "text", "atoms": atoms[pos:s]})
        latex = linear_to_latex(atoms_to_linear(atoms, s, e))
        segs.append({"kind": "math", "atoms": atoms[s:e], "latex": latex})
        pos = e
    if pos < len(atoms):
        segs.append({"kind": "text", "atoms": atoms[pos:]})
    return _merge_text(segs)


def _rpr_key(rpr):
    if rpr is None:
        return None
    return tuple(sorted((c.tag, tuple(sorted((k, v) for k, v in c.attrib.items())))
                        for c in rpr))


def _merge_text(segs: list[dict]) -> list[dict]:
    """合并相邻 text 段，并按 rPr 变化再切分，尽量保留原格式。"""
    out: list[dict] = []
    for seg in segs:
        if seg["kind"] != "text":
            out.append(seg)
            continue
        cur: list[dict] = []
        cur_key = object()
        for a in seg["atoms"]:
            k = _rpr_key(a["rpr"])
            if cur and k != cur_key:
                out.append({"kind": "text", "atoms": cur})
                cur = []
            cur_key = k
            cur.append(a)
        if cur:
            out.append({"kind": "text", "atoms": cur})
    return out
