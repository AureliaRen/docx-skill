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
POW_UNIT_RE = re.compile(
    rf"(\^\{{[-\u2212+]?\d+\}})\s*((?:{_UNIT_ALT})(?:\s*[·⋅]\s*(?:{_UNIT_ALT}))*)"
    rf"(?![_^{{A-Za-z0-9])"
)

# 关系符两边都要有「非运算符」的东西（挡掉 b=、=R、Uoc →，但允许 =∫、=Δ）
_BOTH_SIDES = re.compile(
    r"[^\s=≈≤≥≠∝±<>]\s*[=≈≤≥≠∝±<>]\s*[−+\-]?\s*[^\s=≈≤≥≠∝±<>]")

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
        ok = (
            len(text) >= 2
            and not _looks_like_date(text)
            and alnum >= 1
            and (has_op or has_strong)
            and bool(_BOTH_SIDES.search(text))
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
# 命令后的「可读空格」哨兵：分词时忽略（这样 \partial z 仍算邻接因子），
# 最终输出时再换成真正的空格（否则 \partialx 会被吃成一个命令）。
SP = "\x01"


def _escape_unicode(s: str) -> str:
    for comb, cmd in COMBINING.items():
        s = re.sub(r"(\\[A-Za-z]+|[A-Za-z0-9\)\]}])" + re.escape(comb),
                   lambda m, c=cmd: "\\" + c + "{" + m.group(1) + "}", s)
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
_OP_SPLIT = set("+=<>±∓,;|/÷")


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
            toks[ll:rr + 1] = [("atom", r"\frac{" + L + "}{" + R + "}")]
            changed = True
            break
    return "".join(t for _, t in toks)


def _unitify(s: str) -> str:
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


def linear_to_latex(s: str) -> str:
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

    def _protect(m):
        protected.append(m.group(0))
        return f"\x00{len(protected) - 1}\x00"

    s = re.sub(r"\\text\{[^{}]*\}", _protect, s)
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
    s = _fracify(s)
    # 6) 还原占位符
    s = re.sub(r"\x00(\d+)\x00", lambda m: protected[int(m.group(1))], s)
    s = re.sub(r"\x02(\d+)\x02", lambda m: envs[int(m.group(1))], s)
    s = s.replace(SP, " ")                 # 命令后的可读空格（此时才落成真空格）
    s = re.sub(r" {2,}", " ", s).strip()
    s = re.sub(r"\^([^\s{])", r"^{\1}", s)
    s = re.sub(r"_([^\s{])", r"_{\1}", s)
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
