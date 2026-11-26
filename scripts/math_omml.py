r"""LaTeX → OMML 共享引擎：latex2mathml + 微软 MML2OMML.XSL。

被 `insert_formula.py`（按占位/位置插入单条公式）与 `docx_mathify.py`
（把整篇文档里的线性公式批量转成原生公式）共用，避免两处各写一份转换逻辑。

依赖：
    python -m pip install latex2mathml      # 纯 Python、零依赖、离线可用
    MML2OMML.XSL 随 Microsoft Office 安装，自动探测；可用环境变量
    MML2OMML_XSL 指定路径。
"""
from __future__ import annotations

# --- repo bootstrap: make sibling imports work from any cwd ---------------
import _bootstrap  # noqa: F401 — UTF-8 stdio on Windows
# --------------------------------------------------------------------------

import os
import re
from pathlib import Path

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 内容模型：这些容器只接受「数学项」，不接受嵌套 <m:oMath>
ARG_TAGS = {"num", "den", "e", "sub", "sup", "deg", "fName", "lim"}


# ── 异常类型分层（P28-A） ────────────────────────────────────────────────────
# 为什么要分层：`insert_formula` 早先只写 `except (RuntimeError, FileNotFoundError)`
# 一刀切退 2，于是「公式内容转不了」和「你参数写错了」共用一个码，而 D-024 对 2 的
# 判据是「拒绝先于任何读写 + stderr 恰好一行」—— 守卫的四段多行文案直接把后半条破掉。
# 判据**不能**靠正则匹配文案（文案改一次就静默打歪），所以改成结构化信号：抛点按
# 性质落到下面的子类，调用方按类型分派退出码。
#
# 全部继承 RuntimeError：既有 `except RuntimeError` / `pytest.raises(RuntimeError)`
# 的行为一字不变（`docx_mathify` / `docx_create` 的兜底分支正是这么写的）。
class MathOmmlError(RuntimeError):
    """LaTeX→OMML 全部可预期失败的基类；本身不直接抛，只用于类型归组与兜底识别。"""


class LatexGuardError(MathOmmlError):
    """层 1 硬拒 / 层 2 残留（源文本含 OMML 表达不了的控制序列）→ `insert_formula` 退 1。"""


class LatexSyntaxError(MathOmmlError):
    """`latex2mathml` 语法错 / `MML2OMML.XSL` 应用失败（公式本身过不去）→ `insert_formula` 退 1。"""


class OMMLStructureError(MathOmmlError):
    """`assert_legal` 判四类结构非法 / XSLT 没产出 OMML 根元素 → `insert_formula` 退 1。"""


class MathEnvError(MathOmmlError):
    """转换环境不可用（缺 `latex2mathml` / `MML2OMML.XSL` 找不到或加载不了）→ `insert_formula` 退 2。"""

# ── 公式字体 ─────────────────────────────────────────────────────────────────
# Word 的数学区默认落 Cambria Math。主人要求默认 Times New Roman。
# 做法：给每个 <m:r> 补 <w:rPr><w:rFonts .../></w:rPr>。
# 用环境变量 MATH_FONT 可整体覆盖（如 "Cambria Math" 还原 Word 默认观感）。
DEFAULT_MATH_FONT = "Times New Roman"


def math_font() -> str | None:
    """当前生效的公式字体；返回 None 表示不注入（用 Word 默认 Cambria Math）。"""
    env = os.environ.get("MATH_FONT")
    if env is not None:
        env = env.strip()
        return env or None
    return DEFAULT_MATH_FONT


def _w(tag: str) -> str:
    return f"{{{W_NS}}}{tag}"


def _m(tag: str) -> str:
    return f"{{{M_NS}}}{tag}"


def apply_math_font(omml, font: str | None = None):
    """给 OMML 里每个 <m:r> 注入 w:rPr/w:rFonts，让公式用指定字体渲染。

    只改字体族，不动 <m:rPr>/<m:sty>，所以斜体等数学样式保持 Word 默认。
    ``font`` 为 None 时取 ``math_font()``；再为 None 则原样返回。
    """
    from lxml import etree

    if font is None:
        font = math_font()
    if not font:
        return omml

    # 确保 w 命名空间绑定到 "w" 前缀；否则 lxml 会生成 ns0:rPr 这种前缀，
    # 虽然 XML 语义等价，但写进 document.xml 后和人手写的 OOXML 不一致，
    # 也会让下游按字符串匹配 "w:rPr" 的脚本失效。
    nsmap = dict(omml.nsmap or {})
    if nsmap.get("w") != W_NS:
        nsmap["w"] = W_NS
        new_root = etree.Element(omml.tag, nsmap=nsmap)
        for k, v in omml.attrib.items():
            new_root.set(k, v)
        for child in list(omml):
            new_root.append(child)
        omml = new_root

    for run in omml.iter(_m("r")):
        # CT_R 顺序：m:rPr? → w:rPr? → (m:t | …)，w:rPr 必须在 m:t 之前
        rpr = run.find(_w("rPr"))
        if rpr is None:
            rpr = etree.SubElement(run, _w("rPr"))
            # SubElement 追加到末尾，挪到 m:t 前面
            run.remove(rpr)
            idx = 1 if run.find(_m("rPr")) is not None else 0
            run.insert(idx, rpr)
        # w:rPr 子元素顺序：rStyle → rFonts → b → i → …
        fonts = rpr.find(_w("rFonts"))
        if fonts is None:
            fonts = etree.Element(_w("rFonts"))
            rpr.insert(0, fonts)
        # ascii/hAnsi 管拉丁字母与数字，cs 管复杂文种，eastAsia 管中文
        fonts.set(_w("ascii"), font)
        fonts.set(_w("hAnsi"), font)
        fonts.set(_w("cs"), font)
        # 公式里的中文（\text{中文}）也跟随，避免回落到主题字体
        fonts.set(_w("eastAsia"), font)

    return omml


# ── 直立补全（m:sty） ────────────────────────────────────────────────────────
# Word 的数学区里字母默认**斜体**（变量），数字 / 运算符 / 函数名才该直立。
# latex2mathml 只在少数场合自带 m:sty="p"（\sin / \log / \mathrm{...}），
# MML2OMML.XSL 不补这一层 —— 于是 \max\{x,y\} 在 Word 里是一串斜体的 max{x,y}。
#
# **为什么必须拆 run，而不是给 run 补个属性**：m:sty 是 run 级属性，而一个 m:r
# 经常同时装着两类字符（\frac{dy}{dx} 的分子是单个 run "dy"、分母是单个 run "dx"）。
# 给 "dx" 补 p 会把 x 一起变直立 —— 更错。只能先按语义分块，再让每块各拿各的 m:sty。
#
# 白名单里**刻意不放 "E"**：工程 / 物理文档里的 E 几乎总是电动势或能量（斜体变量），
# 把 E 一律变直立会大面积误伤；期望算子 E[·] 在中文交付物里极少出现，不值得用它的
# 正确性去换全篇 E 的错误。
UPRIGHT_FUNCS = (
    # 长名在前（最长前缀优先），匹配时再按长度降序兜一遍
    "arctan", "arcsin", "arccos", "argmax", "argmin",
    "sinh", "cosh", "tanh", "coth",
    "sin", "cos", "tan", "cot", "sec", "csc",
    "log", "ln", "lg", "lim", "max", "min", "sup", "inf",
    "det", "exp", "gcd", "lcm", "dim", "ker", "deg", "mod",
    "Pr", "Var", "Cov",
)

# 逐字符扫描时按长度降序试白名单：'sinh' 必须先于 'sin' 命中，否则 'sinh'
# 会被切成 'sin' + 'h'（h 变斜体）。
_UPRIGHT_FUNCS_BY_LEN = tuple(sorted(UPRIGHT_FUNCS, key=len, reverse=True))


def _local(tag) -> str:
    """去掉 {ns} 前缀取本地名；非字符串标签（注释等）原样返回。"""
    if isinstance(tag, str) and tag.startswith("{"):
        return tag.rsplit("}", 1)[-1]
    return tag


def classify_run_text(text: str) -> list[tuple[str, str | None]]:
    r"""把一个 m:r 的文本切成 ``[(chunk, "p" | None)]``。

    "p" = 该块要直立（数字、运算符、函数名），None = 保持 Word 默认的数学斜体
    （变量）。

    判据（与 work/p14-round/PLAN-P14.md §3 的实测一致）：

    * 空白 / 不换行空格（U+00A0，``\ `` 控制空格的产物）→ None。它本来就不需要
      sty；上游已经给过 p 的 run 会先被整 run 跳过，走不到这里；
    * 数字串（含小数点与逗号连写，3.14 / 1,2）→ p；
    * 非字母非数字（运算符、括号、花括号、逗号、等号、U+2212 减号、乘号、撇号…）→ p；
    * 字母串：先试函数名白名单，**命中条件是函数名后不是小写字母**（后接 { / ( /
      空格 / 大写 / 结尾都算命中；后接小写字母说明那是变量名，如 maxx）；整段函数名
      算一块 p；
    * 微分算子：d 紧跟字母 → ("d", "p")，后续字母照常走（dx → 'd'(p) + 'x'(None)）；
    * 其余单个字母 → None。
    """
    out: list[tuple[str, str | None]] = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch.isspace() or ch == "\u00a0":
            out.append((ch, None))
            i += 1
            continue
        if ch.isdigit() or ch in ".,":
            j = i
            while j < n and (text[j].isdigit() or text[j] in ".,"):
                j += 1
            out.append((text[i:j], "p"))
            i = j
            continue
        if not ch.isalpha():
            out.append((ch, "p"))
            i += 1
            continue
        hit = None
        for func in _UPRIGHT_FUNCS_BY_LEN:
            if text.startswith(func, i):
                nxt = text[i + len(func): i + len(func) + 1]
                # 后接小写字母 = 变量名的一部分（maxx），不认
                if not (nxt and nxt.islower()):
                    hit = func
                    break
        if hit:
            out.append((hit, "p"))
            i += len(hit)
            continue
        if ch == "d" and i + 1 < n and text[i + 1].isalpha():
            out.append(("d", "p"))
            i += 1
            continue
        out.append((ch, None))
        i += 1
    return out


def split_runs_for_upright(omml, stats: dict | None = None) -> int:
    r"""按字符分块拆 m:r，给该直立的块补 ``m:rPr/m:sty val="p"``。

    返回被**拆分（重建）**的 run 数。stats 非 None 时累进
    ``stats["upright_split"]``（拆了几个 run）与 ``stats["upright_runs"]``
    （共补了几段 m:sty）。

    三种 run 一律整 run 跳过，不动：

    * 带 ``m:rPr/m:nor`` 的 —— \text{...} 的产物，nor 本身已强制直立，再拆会弄坏；
    * 已带非空 ``m:rPr/m:sty`` 的 —— 上游（\sin / \mathrm{Var}）已经定过，尊重原值；
      重复处理只会把已有样式覆盖掉；
    * 所有块都不需要直立的（纯变量名，如 xy）。

    新增的每段都 deepcopy 原 run 的 w:rPr：字体是在 ``apply_math_font`` 阶段注入的，
    重建 run 时不带上就会回落 Cambria Math —— 一条公式里字体半截 TNR 半截 Cambria，
    比整条都错还难查。所以本函数必须在 ``apply_math_font`` **之后**跑。
    """
    from copy import deepcopy

    from lxml import etree

    split = 0
    upright = 0
    for run in list(omml.iter(_m("r"))):
        rpr = run.find(_m("rPr"))
        if rpr is not None:
            if rpr.find(_m("nor")) is not None:
                continue
            sty = rpr.find(_m("sty"))
            if sty is not None and sty.get(_m("val")):
                continue

        t = run.find(_m("t"))
        if t is None or not t.text:
            continue
        parts = classify_run_text(t.text)
        if all(kind is None for _chunk, kind in parts):
            continue

        parent = run.getparent()
        if parent is None:
            continue
        idx = parent.index(run)
        w_rpr = run.find(_w("rPr"))

        built = []
        for chunk, kind in parts:
            seg = etree.Element(_m("r"))
            if w_rpr is not None:
                seg.append(deepcopy(w_rpr))
            nt = etree.SubElement(seg, _m("t"))
            nt.text = chunk
            # 原 m:t 可能带 xml:space="preserve"（含首尾空格的段），照抄
            nt.attrib.update(t.attrib)
            if kind == "p":
                # CT_R 的内容模型是 m:rPr? → w:rPr? → m:t，m:rPr 必须首子元素
                mrpr = etree.Element(_m("rPr"))
                etree.SubElement(mrpr, _m("sty")).set(_m("val"), "p")
                seg.insert(0, mrpr)
                upright += 1
            built.append(seg)

        parent.remove(run)
        for off, seg in enumerate(built):
            parent.insert(idx + off, seg)
        split += 1

    if stats is not None:
        stats["upright_split"] = stats.get("upright_split", 0) + split
        stats["upright_runs"] = stats.get("upright_runs", 0) + upright
    return split


# ── cases 左对齐 ─────────────────────────────────────────────────────────────
# OMML 里 cases 与 pmatrix / bmatrix / vmatrix / matrix **完全同构**：都是 m:m
# + baseJc=center + plcHide=on + mcs/mc/mcPr/{count,mcJc=center}。光看 OMML 分不出
# 谁是 cases，所以只能在 MathML 层先判，再按文档序对位：
#
# * 判据在 MathML：\begin{cases} 的每个 mtd 都带 columnalign="left"，矩阵环境的
#   mtd 不带任何 columnalign；
# * 映射规则（实测）：只有「列数 ≥ 2」的 mtable 才产出 m:m，且按文档序 1:1；
# * 长度对不上就整段放弃 —— 宁可保持 center（观感略差），也不赌一个可能错的位移。
MATHML_NS = "http://www.w3.org/1998/Math/MathML"

# 补建 m:mPr / m:mcPr 时用的子元素顺序（CT_MPr / CT_MCPr 的内容模型）。
_MPR_ORDER = ("baseJc", "plcHide", "rSpRule", "cGpRule", "rSp", "cSp", "cGp",
              "mcs", "ctrlPr")
_MCPR_ORDER = ("count", "mcJc")


def _insert_ordered(parent, child, order: tuple[str, ...]) -> None:
    """按 order 把 child 插到第一个「排在它后面」的兄弟之前，没有就追加。"""
    pos = order.index(_local(child.tag))
    for i, existing in enumerate(parent):
        name = _local(existing.tag)
        if name in order and order.index(name) > pos:
            parent.insert(i, child)
            return
    parent.append(child)


def _force_mc_jc_left(m) -> bool:
    """把 m:m 的 mcs/mc/mcPr/mcJc@val 设成 left（缺哪层补哪层）。"""
    from lxml import etree

    mpr = m.find(_m("mPr"))
    if mpr is None:
        mpr = etree.Element(_m("mPr"))
        m.insert(0, mpr)
    mcs = mpr.find(_m("mcs"))
    if mcs is None:
        mcs = etree.Element(_m("mcs"))
        _insert_ordered(mpr, mcs, _MPR_ORDER)
    mc = mcs.find(_m("mc"))
    if mc is None:
        mc = etree.SubElement(mcs, _m("mc"))
    mcpr = mc.find(_m("mcPr"))
    if mcpr is None:
        mcpr = etree.Element(_m("mcPr"))
        mc.insert(0, mcpr)
    jc = mcpr.find(_m("mcJc"))
    if jc is None:
        jc = etree.Element(_m("mcJc"))
        _insert_ordered(mcpr, jc, _MCPR_ORDER)
    if jc.get(_m("val")) == "left":
        return False
    jc.set(_m("val"), "left")
    return True


def align_cases_left(mathml_root, omml, stats: dict | None = None) -> int:
    r"""把 cases 环境的列对齐改成左对齐，返回改了几个 m:m。

    做法见本节顶部注释：先收集 MathML 里「列数 ≥ 2」的 mtable（文档序）与它是否
    cases 样（所有 mtd 的 columnalign == "left"），再与 OMML 里的 m:m（文档序）
    逐位对。只有 cases 样的那个才改。

    两个列表长度不等 → **什么都不改**，把原因记进
    ``stats["cases_align_skipped"]``（形如 "mtable 2 vs m:m 3"），返回 0。
    这是刻意保守：错位改一格，会把矩阵变成左对齐，而且没有任何测试能提前发现。
    """
    tables: list[bool] = []
    for tbl in mathml_root.iter("{%s}mtable" % MATHML_NS):
        rows = [r for r in tbl if _local(r.tag) in ("mtr", "mlabeledtr")]
        cells = [c for r in rows for c in r if _local(c.tag) == "mtd"]
        ncols = max((sum(1 for c in r if _local(c.tag) == "mtd")
                     for r in rows), default=0)
        if ncols < 2:
            continue                     # 1 列环境不产 m:m，不参与对位
        tables.append(bool(cells)
                      and all(c.get("columnalign") == "left" for c in cells))

    mats = list(omml.iter(_m("m")))
    if len(tables) != len(mats):
        if stats is not None:
            stats.setdefault("cases_align_skipped", []).append(
                "mtable %d vs m:m %d" % (len(tables), len(mats)))
        return 0

    changed = 0
    for cases_like, m in zip(tables, mats):
        if cases_like and _force_mc_jc_left(m):
            changed += 1
    if stats is not None:
        stats["cases_aligned"] = stats.get("cases_aligned", 0) + changed
    return changed

def wrap_leading_unary_in_delim(mathml_root, stats: dict | None = None) -> int:
    r"""把定界组里「组首一元正负号 + 紧随其后的那一项」包成一个 mrow。

    ## 为什么在 MathML 层修，而不是修 XSL 的产物

    定界组经 ``latex2mathml`` → ``MML2OMML.XSL`` 之后，XSL 会把组首的一元
    ``-`` / ``+`` / ``±`` / ``∓`` 当成**槽分隔符**（``m:sepChr``），产出
    ``<m:d><m:dPr><m:sepChr m:val="−"/></m:dPr><m:e/><m:e>…</m:e></m:d>`` ——
    第一个 ``m:e`` 是空的，Word 把它画成一个空方块。这不是本仓库任何一行的错，
    已实测：把 ``latex2mathml`` 的 MathML 直接喂 ``math_omml.xslt()`` 就复现，
    连 skill 代码都不经过。

    所以修在**上游**（XML 树进 XSLT 之前），而不是去重建 XSL 的 OMML 产物：
    OMML 侧的定界组是 ``m:d`` + 多个 ``m:e``，要在那一层合并得先判断「谁是槽、
    谁是真内容」，还得重建 run 与字体注入，动一处错一片；而在 MathML 侧这只是一
    次 ``mrow`` 重新分组，XSLT 之后自然得到**单个**非空 ``m:e``，Word 不再画方块。

    ## 怎么认出「一个定界组」（**不是**看 ``fence`` 属性）

    ``\left( … \right)`` 产的 ``<mo>`` 带 ``fence="true"``，但**普通括号不带**
    （只有 ``stretchy="false"``）。而正文里 ``(-\frac{y}{x})`` 这种普通括号写法
    同样会踩这个坑（实测同样 1 个空 ``m:e`` + ``sepChr="−"``），所以判据不能挂在
    ``fence`` 上，否则最常见的一半完全修不到。

    改用**平衡深度**：该 ``mrow`` 的首元素是开定界符、末元素是对应的闭定界符，
    且从首元素往后数到（不含）末元素为止，深度**从不归零**。归零就说明中途已经
    闭合过，那不是一个组而是并排的几个（``(-x)(-y)`` 就是这样，它本来也不坏，
    但语义上不能当成一个组去包）。

    ## 命中后怎么改

    去掉首尾定界符之后的 body 长度 ≥ 2，且 ``body[0]`` 是 ``<mo>``、其文本
    ``.strip()`` 属于 ``−``(U+2212) / ``-`` / ``+`` / ``±`` / ``∓``（第一个是
    U+2212 MINUS SIGN，不是 ASCII 连字符，两者都要收），命中则把 ``body[0]`` 与
    ``body[1]`` 一起包进新的 ``<mrow>``，插回 ``body[0]`` 原来的位置。

    **``\left(a+b\right)`` 这类正常分隔符不动**：组首不是一元正负号，一条都命中
    不了。``(-x+y)`` / ``(-\frac{a}{b}+1)`` / ``(-x,\ y)`` 这三条**会**命中 1 次
    wrap（组首确实是一元负号），但 body[1] 只把「−x」包起来，后面那个 ``+`` /
    ``,`` 仍在组内 —— 它们本来就产 0 个空槽，wrap 之后 ``sepChr`` 由**空**变成
    组内真正的 ``"+"`` / ``","``、槽由一个变两个，Word 渲染等价（测试专门钉这三
    条，钉的是「空槽数为 0 且 sepChr 是变后的真值」）。

    **已知不修**：``\left(\right)`` / ``()`` 这种空组，组内 0 个项、没有
    ``body[1]`` 可包，仍会留 1 个空 ``m:e`` —— 那属于「用户写了个空括号」，
    交给 :func:`assert_legal` 的新检查报错，不替它编造内容。

    返回包装次数；``stats`` 非 None 时累进 ``stats["delim_wraps"]``（整型累加，
    与 ``upright_split`` / ``cases_aligned`` 同风格）。
    """
    from lxml import etree

    mrow_tag = "{%s}mrow" % MATHML_NS
    mo_tag = "{%s}mo" % MATHML_NS
    pairs = {"(": ")", "[": "]", "{": "}", "|": "|",
             "‖": "‖", "⟨": "⟩",
             "⌈": "⌉", "⌊": "⌋"}
    opens = set(pairs)
    closes = set(pairs.values())
    unary = {"−", "-", "+", "±", "∓"}

    def is_single_group(kids) -> bool:
        """首尾是一对匹配定界符、中间深度不归零 → 这是一个完整的定界组。"""
        if len(kids) < 3:
            return False
        first, last = kids[0], kids[-1]
        if first.tag != mo_tag or last.tag != mo_tag:
            return False
        o = (first.text or "").strip()
        c = (last.text or "").strip()
        if o not in opens or pairs[o] != c:
            return False
        depth = 0
        for k in kids[:-1]:
            if k.tag != mo_tag:
                continue
            t = (k.text or "").strip()
            if t in opens:
                depth += 1
            elif t in closes:
                depth -= 1
            if depth == 0:
                # 中途闭合：并排的好几个组（(-x)(-y)），不是一个组
                return False
        return True

    wrapped = 0
    for mrow in list(mathml_root.iter(mrow_tag)):
        kids = [c for c in mrow if isinstance(c.tag, str)]
        if not is_single_group(kids):
            continue
        body = kids[1:-1]
        if len(body) < 2:
            continue
        head = body[0]
        if head.tag != mo_tag:
            continue
        if (head.text or "").strip() not in unary:
            continue

        group = etree.Element(mrow_tag)
        mrow.insert(mrow.index(head), group)
        group.append(head)          # append 会自动从旧的 mrow 摘除
        group.append(body[1])
        wrapped += 1

    if stats is not None:
        stats["delim_wraps"] = stats.get("delim_wraps", 0) + wrapped
    return wrapped

def find_xslt() -> Path:
    """定位 MML2OMML.XSL。"""
    env = os.environ.get("MML2OMML_XSL")
    if env:
        p = Path(env)
        if p.is_file():
            return p
        raise FileNotFoundError(f"MML2OMML_XSL={env} 不是文件")

    roots = [
        os.environ.get("ProgramFiles"),
        os.environ.get("ProgramFiles(x86)"),
        os.environ.get("LOCALAPPDATA"),
        r"C:\Program Files",
        r"C:\Program Files (x86)",
    ]
    patterns = (
        "Microsoft Office/**/MML2OMML.XSL",
        "Microsoft Office/MML2OMML.XSL",
        "*/*/Office*/MML2OMML.XSL",
        "Kingsoft/**/MML2OMML.XSL",
    )
    for root in roots:
        if not root or not Path(root).is_dir():
            continue
        for pat in patterns:
            try:
                hits = sorted(Path(root).glob(pat))
            except (OSError, ValueError):
                continue
            if hits:
                return hits[-1]
    raise FileNotFoundError(
        "找不到 MML2OMML.XSL（随 Microsoft Office/Word 安装）。\n"
        "  1) 确认已装 Office，且存在 <Office目录>\\root\\Office16\\MML2OMML.XSL\n"
        "  2) 或用环境变量指定：$env:MML2OMML_XSL = '<路径>'"
    )


_XSLT = None


def xslt():
    r"""编译并缓存 XSLT → (transform, path)。

    环境类失败一律转成 :class:`MathEnvError`（``insert_formula`` 退 2，stderr 单行）：
    ``MML2OMML.XSL`` 存在但**不是合法 XML** 时，裸 ``etree.parse`` 会掀
    ``lxml.etree.XMLSyntaxError`` 的 29 行 traceback（P28 探针 5 实测）。

    实现选择：**一个 try 同时包住** ``etree.parse`` 与 ``etree.XSLT``（两处失败性质
    相同 —— 都是「这份 XSL 加载不了」；分成两个 try 只会重复同一段转换逻辑）。
    消息里带原始 lxml 原因，并压成**单行**（2 档要求 stderr 恰好一行）。

    ⚠️ ``find_xslt()`` 的 ``FileNotFoundError`` **不进**这个 try：它是「文件不存在」，
    探针 4 已合规（rc=2 + 单行），吞进来会改掉既有语义与文案。
    """
    global _XSLT
    if _XSLT is None:
        from lxml import etree
        path = find_xslt()          # FileNotFoundError 原样冒出，不包进下面
        try:
            transform = etree.XSLT(etree.parse(str(path)))
        except (etree.XMLSyntaxError, etree.XSLTParseError, OSError) as exc:
            # 只取首行：lxml 的多行报错会把上下文一起打出来，破「stderr 恰好一行」。
            first = (str(exc).splitlines() or [""])[0]
            raise MathEnvError(
                f"MML2OMML.XSL 无法加载（{path}）: {type(exc).__name__}: {first}"
            ) from exc
        _XSLT = (transform, str(path))
    return _XSLT


def keep_spacing(latex: str) -> str:
    r"""间距命令会被 XSLT 丢掉（普通空格也会），改写为 ``\ ``（控制空格 → NBSP）。"""
    import re
    return re.sub(r"\\(?:qquad|quad|,|;|:|!)", r"\\ ", latex)


# ── 控制序列守卫（P27-A）：LaTeX 专属命令硬拒 + 残留兜底 ─────────────────────
#
# ## 为什么需要这一层
#
# latex2mathml 对**不认识的控制序列不报错**，而是把它们当字面量塞进 <mi>
# → MML2OMML.XSL 原样搬进 <m:t>。于是交付物的公式里真的印着 \eqno、\textwidth，
# 而整条链路四个环节全部沉默：转换不抛、docx_mathify 的 stats["failed"] 只数
# 异常、退出码只看那个计数、assert_legal 只查结构不看文本。取证见
# docs/reports/P27-PROBE-CONTROLSEQ.md（48 条矩阵：leak 25 / error 1 / dropped 1
# / ok 21）。
#
# ## 两层，各自不可省
#
# 1. **前置硬拒**（find_unsupported_commands）—— 治因。表里穷举「LaTeX 排版
#    专属、OMML 里没有对应物」的那批命令，命中就在**转换之前**抛 RuntimeError，
#    并逐条给出替代写法。代价是表永远不全：\zzz 这种用户瞎写的命令不可能预先列出。
# 2. **后置残留检测**（find_latex_residue）—— 兜底。扫成品 OMML 的每个 m:t，
#    只要某段文本**整段**就是一条控制序列就拒收。这一层不看命令名，所以任意未知
#    命令都跑不掉 —— 去掉它就等于把 \zzz 放行。
#
# ## 上下文档位是关键，不是锦上添花
#
# \nolimits / \middle 这类命令在特定位置是**合法**的：实测
# \sum\nolimits_{i=1}^{n} a_i 正常、x\nolimits 泄漏；\left(a \middle| b \right)
# 正常、裸 a \middle| b 报错。按命令名一刀切必然误伤正常公式，所以它们不进
# 黑名单，而是按位置判定（见 CONTEXT_RULES）。

# \limits / \nolimits 只有**紧跟**大算符时才是合法修饰符。
BIG_OPERATORS = (
    "sum", "prod", "coprod", "int", "iint", "iiint", "oint",
    "bigcup", "bigcap", "bigvee", "bigwedge", "bigoplus", "bigotimes",
    "bigodot", "biguplus", "lim", "limsup", "liminf", "sup", "inf",
    "max", "min", "det", "gcd", "Pr",
)

# 每条命令一行中文替代写法。**只收实测会出事的**：P27-PROBE-CONTROLSEQ \② 的 25
# 条泄漏 + work/p27a/probe*.py 复核过的同族扩展；**本轮（P29-A）另加 6 条「静默
# 失真」族**（\rule / \hspace / \hskip / \kern / \mspace / \enspace）—— 它们不
# 泄漏反斜杠，但经 latex2mathml 变成 <mspace> 之后，MML2OMML.XSL（全文无 mspace
# 模板）会把整条公式吞成**零子元素的 <m:oMath/>**，Word 里什么都不显示。取证见
# docs/reports/P29-A-GUARD.md \1（\rule{10pt}{2pt} / \hspace{1em} 实测 0 子元素，
# \hspace*{1em} 实测泄漏 <m:t>1em</m:t>）。
# \def / \newcommand **也不进**：latex2mathml 会展开简单宏
# （\def\foo{1}\foo 实测正常），硬拒会误伤；它们由 CONTEXT_RULES 的 macro_use
# 判据处理（定义了却没用 = 静默产出空公式，那种才拒）。
LATEX_ONLY_COMMANDS: dict[str, str] = {
    # —— 编号 / 引用：LaTeX 的 \label-\ref 体系，OMML 里没有对应物 ——
    # ⚠️ 这三条（label / eqref / ref）**在整篇链路里是死代码，在单条入口上是唯一防线**：
    #    整篇走 docx_mathify 时，`_mark_commands()`（docx_mathify.py:498）在 `_analyze`
    #    之前就把它们摘成 special atom 做编号，于是 \label{} 根本到不了
    #    latex_to_omml_lxml，这一层拦不到、也不需要拦；而单条入口（insert_formula /
    #    docx_create 的 latex block）**没有**那层预处理，\label 会一路走到
    #    latex2mathml 变成字面量。所以保留，但说清它是哪条入口的防线。
    "eqno": r"去掉 \eqno：编号改用 \tag{n}，本仓库的编号链路会把它落成正文里的 (n)",
    "nonumber": r"去掉 \nonumber：编不编号由编号段 / \tag 决定，不需要逐行压制",
    "notag": r"去掉 \notag：与 \nonumber 同族，OMML 里没有「这一行不编号」的位",
    "label": r"去掉 \label{}：交叉引用交给 docx_crossref（书签 + REF 域），公式里不留标签",
    "eqref": r"改为手写编号，或交给 docx_crossref 生成 REF 域；公式里不写 \eqref{}",
    "ref": r"改为手写编号，或交给 docx_crossref 生成 REF 域；公式里不写 \ref{}",
    # —— 版面尺寸：Word 里由段落 / 表格 / 页面设置负责 ——
    "textwidth": r"改为具体长度（\textwidth 换成绝对值），或把公式放进整行宽的段落",
    "textheight": "版面高度交给页面设置（docx_create 的 page 段），公式里不要引用",
    "columnwidth": "改为具体长度；栏宽由 Word 的分栏设置决定",
    "linewidth": "改为具体长度；行宽由 Word 的段落 / 表格设置决定",
    "baselineskip": "行距交给段落格式（line_spacing），公式里不要引用",
    "setlength": r"去掉 \setlength：长度由 docx_create / docx_edit 的格式参数设定",
    # —— 垂直 / 水平间距：OMML 无对应物，间距交给段落 ——
    "vspace": r"去掉 \vspace：段间距用段落格式的 space_before / space_after",
    "hfill": r"去掉 \hfill：要右对齐就用右对齐段落，公式里不写",
    "vfill": r"去掉 \vfill：垂直填充由页面 / 段落布局负责",
    "hfil": r"去掉 \hfil：水平填充由对齐方式负责",
    "smallskip": r"去掉 \smallskip：段间距用段落格式设定",
    "medskip": r"去掉 \medskip：段间距用段落格式设定",
    "bigskip": r"去掉 \bigskip：段间距用段落格式设定",
    "bigskipamount": r"去掉 \bigskipamount：段间距用段落格式设定",
    # —— 断行 / 断页：Word 里是段落与分页符，不是公式内容 ——
    "cr": r"去掉 \cr：矩阵 / 对齐环境里换行用双反斜杠，公式外的换行用新段落",
    "newline": r"去掉 \newline：公式内换行用双反斜杠（或拆成两条公式）",
    "linebreak": r"去掉 \linebreak：换行交给段落",
    "nolinebreak": r"去掉 \nolinebreak：断行控制交给段落，或用不换行空格的写法",
    "pagebreak": r"去掉 \pagebreak：分页用 docx_edit 的分页符，不写在公式里",
    "newpage": r"去掉 \newpage：分页用 docx_edit 的分页符",
    "clearpage": r"去掉 \clearpage：分页用 docx_edit 的分页符",
    "allowbreak": r"去掉 \allowbreak：断行点由排版引擎决定，OMML 无此概念",
    "break": r"去掉 \break：断行交给段落",
    "displaybreak": r"去掉 \displaybreak：多行公式的断页由段落分页控制",
    "allowdisplaybreaks": r"去掉 \allowdisplaybreaks：多行公式的断页由段落分页控制",
    # —— 表格线 / 列合并：Word 表格里没有「公式内的表格线」——
    "hline": r"去掉 \hline：表格加边框用 docx_create / table_width 的表框设置",
    "cline": r"去掉 \cline{}：表格加边框用表框设置",
    "hrule": r"去掉 \hrule：横线用段落边框或表格边框",
    "vrule": r"去掉 \vrule：竖线用表格边框",
    # —— 间距 / 画线：XSL 没有 mspace 模板，这一族会把整条公式吞成空 <m:oMath/> ——
    # 星号变体（\hspace*{1em}）不需要单独一条：_CMD_RE 的 (\*?) 把星号吃掉，
    # 命令名仍是 hspace（P29-A 实测：同样命中本表、同样被拒）。
    "rule": r"去掉 \rule：公式里不画横线 —— 横线用段落边框或表格边框；分数用 \frac；要上划线用 \overline{…}",
    "hspace": r"去掉 \hspace：手写间距在 OMML 里存不住（XSL 没有 mspace 模板，整条公式会被吞空）；要间距就用 \quad（1em）/ \,（3/18em）/ \; / \: / \! / \ （控制空格）—— 这六个是本链路唯一能活下来的间距",
    "hskip": r"去掉 \hskip：它是 TeX 原始的「水平跳距」命令（\hskip 1em），OMML 里没有对应节点、还会把整条公式吞空；改用 \quad / \, / \; / \: / \! / \ 这六个能活的间距",
    "kern": r"去掉 \kern：它是 TeX 原始的「字距微调」命令（\kern 1pt，比 \hskip 少一个可拉伸量），同样会被吞成空公式；微调用 \, / \! ，整档间距用 \quad",
    "mspace": r"去掉 \mspace：它是 MathML 风格的间距写法（\mspace{1mu}），本链路只认 LaTeX 那一族；改用 \quad / \, / \; / \: / \! / \ 控制空格",
    "enspace": r"去掉 \enspace：它是「半个 em」的固定间距（0.5em），XSL 侧同样无处安放；要半档间距就用 \,（3/18em）或 \: ，整档用 \quad",
    "multicolumn": r"去掉 \multicolumn{}：单元格合并用 docx_edit 的 merge_cells",
    # —— 盒子：OMML 无对应结构 ——
    "makebox": r"去掉 \makebox：定宽内容用表格，或改用 \text{}",
    "raisebox": r"去掉 \raisebox：上下移位交给上下标结构",
    # —— 宏 / 版面细调：Word 侧无对应物 ——
    "renewcommand": r"去掉 \renewcommand：latex2mathml 不展开它（实测整串泄漏），改写成可直接转换的命令",
    "settowidth": r"去掉 \settowidth：它量的是 LaTeX 版面长度，OMML 无等价物；要控制宽度请用表格列宽",
    "ignorespaces": r"去掉 \ignorespaces：空格由源文本直接控制",
    "unskip": r"去掉 \unskip：它撤掉的是一个 LaTeX 间距节点，OMML 里没有",
    "protect": r"去掉 \protect：它是 LaTeX 展开控制的修饰符，OMML 无对应物",
    # —— 段落：公式片段里不该出现 ——
    "par": r"去掉 \par：分段用新段落（docx_edit 的 add_paragraph）",
    "noindent": r"去掉 \noindent：首行缩进用段落格式 first_line_indent",
    # —— 文本互操作：数学环境里插正文 ——
    "intertext": r"去掉 \intertext{}：多行公式之间插正文请拆成「公式 + 段落 + 公式」",
    "shortintertext": r"去掉 \shortintertext{}：同 \intertext，拆成公式与段落",
    "qed": r"去掉 \qed：证明结束符用 \square / \Box，或在正文里手写 ∎",
}

# 上下文判据。「按命令名一刀切会误伤」的实测依据见本节顶部注释。
CONTEXT_RULES: dict[str, dict[str, str]] = {
    "limits": {
        "rule": "bigop_prefix",
        # 为什么留这一条（实测，不是「为口径一致」）：
        #   * 顶层位置 \limits 会被 latex2mathml **静默吃掉**（`x\limits` → m:t
        #     ['x']，不泄漏）；
        #   * 但它出现在 \text{…} 里时会跟着正文一起变字面量：
        #     `\text{a\limits}` → m:t ['a\\limits'] ——
        #     而层 2 的整段锚定**看不见这一条**（整段是 `a\limits`，不是 `\limits`）。
        # 所以「\limits 在本链路上从不泄漏」只对顶层成立；有层 2 的盲区就需要层 1。
        # 必要性由 test_limits_rule_is_necessary 用「临时摘掉本条 → 静默泄漏」钉死。
        "reason": r"\limits 只对大算符有意义；非大算符位置会静默失效，在 \text{} 里则会被当字面量印进公式",
        "suggestion": r"去掉 \limits：\limits 只对大算符有意义；在大算符上也只有想让上下限"
                      r"上下排列时才写（写了 \sum\limits 得到 m:limLoc=undOvr，不写默认是"
                      r" subSup；\max\limits 得到 m:limLow）",
    },
    "nolimits": {
        "rule": "bigop_prefix",
        "reason": r"\nolimits 只对大算符有意义；离开大算符它会被当字面量印进公式",
        "suggestion": r"去掉 \nolimits：上下限摆位交给 OMML 结构；要压成行内形态就把上下限写成普通上下标（a_{i}）",
    },
    "middle": {
        "rule": "left_right",
        "reason": r"\middle 只用在 \left … \right 之间",
        "suggestion": r"去掉 \middle：分隔符直接写 \mid / \| / |；要在定界组里当分隔符，就写成 \left( a \middle| b \right)",
    },
    "def": {
        "rule": "macro_use",
        "reason": r"宏定义只有被使用才有意义；未被使用的 \def 会产出一条空公式",
        "suggestion": r"去掉 \def：直接把展开后的内容写进公式（要用宏就确保后面真的引用了它）",
    },
    "newcommand": {
        "rule": "macro_use",
        "reason": r"宏定义只有被使用才有意义；未被使用的 \newcommand 会产出一条空公式",
        "suggestion": r"去掉 \newcommand：直接把展开后的内容写进公式（要用宏就确保后面真的引用了它）",
    },
}

_CMD_RE = re.compile(r"\\([A-Za-z]+)(\*?)")
_RESIDUE_RE = re.compile(r"^\\[A-Za-z]+\*?$")
_LEFT_RIGHT_RE = re.compile(r"\\(left|right)(?![A-Za-z])")
_MACRO_NAME_RE = re.compile(r"[ \t]*\{?[ \t]*\\([A-Za-z]+)")


def _rule_bigop_prefix(latex: str, start: int, _end: int) -> bool:
    r"""前面（跳过空白）紧跟一个大算符 → 放行。

    实测两态：\sum\nolimits_{i=1}^{n} a_i 正常转成上下限结构；x\nolimits 把命令名
    整串印进 m:t。判据是「紧邻」而不是「附近有」：\sum_{i=1}^{n}\nolimits 同样
    泄漏（前缀以右花括号结尾，不是大算符）。
    """
    prefix = latex[:start].rstrip()
    return any(prefix.endswith("\\" + op) for op in BIG_OPERATORS)


def _rule_left_right(latex: str, start: int, _end: int) -> bool:
    r"""位置落在某个未闭合的 \left … \right 区间内 → 放行。

    实测两态：\left(a \middle| b\right) 正常；裸 a \middle| b 由 latex2mathml 抛
    ExtraLeftOrMissingRightError（消息是空串，用户看不出为什么）。
    """
    depth = 0
    for m in _LEFT_RIGHT_RE.finditer(latex, 0, start):
        depth += 1 if m.group(1) == "left" else -1
    return depth > 0


def _rule_macro_use(latex: str, _start: int, end: int) -> bool:
    r"""宏被定义了**且**在后面被引用 → 放行；定义了没用 → 拒。

    实测：\def\foo{1}\foo 与 \newcommand{\x}{1}\x 上游会展开、结果正常；而
    \newcommand{\x}{1} 单独写产出一条空 m:oMath（内容被静默丢掉）。
    """
    rest = latex[end:]
    m = _MACRO_NAME_RE.match(rest)
    if not m:
        return False
    return ("\\" + m.group(1)) in rest[m.end():]


_CONTEXT_PREDICATES = {
    "bigop_prefix": _rule_bigop_prefix,
    "left_right": _rule_left_right,
    "macro_use": _rule_macro_use,
}


def _context_allows(latex: str, m) -> bool | None:
    r"""True 放行 / False 拒 / None 表示这条命令不在上下文表里。"""
    rule = CONTEXT_RULES.get(m.group(1))
    if rule is None:
        return None
    return _CONTEXT_PREDICATES[rule["rule"]](latex, m.start(), m.end())


def find_unsupported_commands(latex: str) -> list[dict]:
    r"""层 1：找出会在 OMML 里变成反斜杠字面量的控制序列。

    返回 [{"command", "star", "position", "reason", "suggestion"}, …]，按源文本
    出现顺序、同名只留第一条。**未知命令不在本函数的射程内** —— \zzz 这种谁都
    想不到的命令由 find_latex_residue 兜。

    带星号（\vspace*{1em}）与带参数（\vspace{1em}）都按同一个命令名识别：星号只
    影响 LaTeX 语义，不影响它会被当字面量印进公式。

    \settowidth{\x}{y} 内层的 \x 属未知命令，**不在这里判**（它自己独占一个 m:t，
    层 2 会抓到）；本函数也不会因为「里面还有未知命令」就放过外层已知命令。
    """
    hits: list[dict] = []
    seen: set[str] = set()
    text = latex or ""
    for m in _CMD_RE.finditer(text):
        name = m.group(1)
        if name in seen:
            continue
        verdict = _context_allows(text, m)
        if verdict is True:
            continue
        if verdict is None and name not in LATEX_ONLY_COMMANDS:
            continue                      # 未知命令：交给后置残留检测
        seen.add(name)
        if verdict is False:
            info = CONTEXT_RULES[name]
            reason, suggestion = info["reason"], info["suggestion"]
        else:
            reason = ("LaTeX 排版专属命令（编号 / 引用 / 版面尺寸 / 断行 / 表格线 / "
                      "间距 / 宏），OMML 里没有对应结构，必然无处安放")
            suggestion = LATEX_ONLY_COMMANDS[name]
        hits.append({
            "command": "\\" + name,
            "star": bool(m.group(2)),
            "position": m.start(),
            "context": verdict is False,
            "reason": reason,
            "suggestion": suggestion,
        })
    return hits


def _guard_hit(stats: dict | None, layer: str, latex: str,
               commands: list[str]) -> None:
    r"""记一次拒收。

    stats 的口径（P27-A 约定，两选一取这个）：命中即累进 stats["latex_guard"]
    （拒收条数），并把明细 append 进 stats["latex_guard_hits"]
    （{"latex", "layer", "commands"}，layer 是 "pre" 或 "residue"）。
    **不回滚此前已累进的排版计数**（upright_split 等）—— 层 2 的检查位置按约定
    必须在 split_runs_for_upright 之后（那里才扫得到最终树），而那时拆 run 的
    计数已经落库。下游要区分「拒收」与「转换失败」看 latex_guard，不要看 failed
    （后者只数异常）。
    """
    if stats is None:
        return
    stats["latex_guard"] = stats.get("latex_guard", 0) + 1
    stats.setdefault("latex_guard_hits", []).append(
        {"latex": latex, "layer": layer, "commands": list(commands)})


def assert_no_unsupported_commands(latex: str, stats: dict | None = None) -> None:
    r"""层 1 的出口：命中就抛 RuntimeError，四段文案（症状 / 成因 / 替代 / 怎么做）。

    调用位置在 latex_to_omml_lxml 里 keep_spacing() **之前** —— 必须扫**原始**
    参数：keep_spacing 会把 \quad / \, 改写成控制空格的形态，扫改写后的串会漏判；
    而且这一层在 latex2mathml 之前就拦下，报错能直接给替代写法，用户不必去读
    上游那个语义不明的语法异常。
    """
    hits = find_unsupported_commands(latex)
    if not hits:
        return
    _guard_hit(stats, "pre", latex, [h["command"] for h in hits])
    # 上下文类命中要带上「为什么这次不行」：同一个 \nolimits 在 \sum 后面是好的，
    # 光给替代写法会让用户以为这条命令一律不许用。
    detail = "\n".join(
        "  - %s%s → %s" % (h["command"], "*" if h["star"] else "",
                          ("（%s）" % h["reason"] if h["context"] else "") + h["suggestion"])
        for h in hits)
    # 编号 / 引用那三条的适用范围要说清：整篇链路里 docx_mathify 的 _mark_commands
    # 会先把它们摘成编号原子（_analyze 之前，docx_mathify.py:498），所以它们到不了
    # 这里 —— 这一层对它们是「单条入口的唯一防线」，不是「整篇链路的拦截点」。
    numbering = sorted({"label", "eqref", "ref"} &
                       {h["command"].lstrip(chr(92)) for h in hits})
    if numbering:
        tail = ("注意：你命中的 %s 属于编号 / 引用类。整篇走 docx_mathify 时，它的 "
                "_mark_commands 会先把这些命令摘成编号原子，不会走到这条守卫；"
                "也就是说这一层拦的是**单条入口**（insert_formula / docx_create 的 "
                "latex block），那里没有那层预处理。整篇链路要改编号请用 "
                "docx_mathify 的编号段 / \\tag。"
                % "、".join("\\" + n for n in numbering))
    else:
        tail = ""
    raise LatexGuardError(
        "LaTeX 含 OMML 表达不了的控制序列（%r）：%s。\n"
        "症状：换成 Word 公式时，命令名可能被当字面量印进公式（例如 \\eqno 原样"
        "出现在公式末尾，回读 document.xml 能搜到以反斜杠开头的 m:t）；也可能直接"
        "抛一个看不懂的语法异常，或者把整条公式静默吞成空的。三种表现都说明这条"
        "命令没有 OMML 对应物。\n"
        "成因：这些命令是 LaTeX 的排版专属（编号 / 引用 / 版面尺寸 / 断行断页 / "
        "表格线 / 宏定义），OMML 里没有对应结构；latex2mathml 对它们要么当普通字符"
        "塞进 <mi>，要么在语法检查上直接失败。\n"
        "替代写法：\n%s\n"
        "怎么做：把源文本里的这些命令按上面的替代写法改掉再转。"
        "%s"
        % (latex, "、".join(h["command"] for h in hits), detail, tail))


def find_latex_residue(omml) -> list[str]:
    r"""层 2：扫成品 OMML 里所有 m:t 的文本，收**以空白为界的词本身是一条命令**的残留。

    返回去重保序的命令名（**带**反斜杠，如 "\zzz"）；干净时返回空列表。

    ## 判据：以空白为界的**词**（token），不是整段 m:t、也不是任意位置

    两个更朴素的判据各自有一个洞，实测都堵不住：

    * **整段锚定**（`^\[A-Za-z]+\*?$` 匹配整条 m:t）**漏文本模式**：
      命令写在 `\text{}` 里时与前后正文**共享一段 m:t** ——
      `\text{a \zzz b}` 转出来是 `m:t` = `a` + NBSP + `\zzz` + NBSP + `b`，
      整段不是单条命令，判据看不见（实测漏，5/5 文本模式全漏）。
    * **任意位置匹配**（`\[A-Za-z]+` 在整段里 search）**误伤正文**：
      `\text{C:\Users\me}` 的 m:t 是 `C:\Users\me` —— 一个 Windows
      路径，任意位置会把 `\Users` 当残留，把一条完全正常的公式判死。

    取中间：**先按空白把 m:t 切成 token，再对每个 token 做整段锚定**。`str.split()`
    恰好按 Unicode 空白切分，**包含 NBSP（U+00A0）** —— 而 latex2mathml 在 `\text{}`
    里正是用 NBSP 填空白，所以 `a\xa0\zzz\xa0b` 切得出 `['a', '\zzz', 'b']`；
    而 `C:\Users\me` 切出来仍是一个 token（开头是 `C:`），不命中。

    ## 实测三判据对照（work/p27a/tok.py，本机 latex2mathml 3.81.1）

    | 判据 | 25 条数学模式泄漏 | 文本模式残留 | 合法公式误报 |
    |---|---|---|---|
    | 整段锚定 | 25/25 | **0/5** | 0 |
    | 任意位置 | 25/25 | 5/5 | **3 条** |
    | **按空白切词** | **25/25** | **5/5** | **0** |

    **只认 m: 命名空间的 m:t**：w:t 里的反斜杠是合法正文（`C:\Users` 那种路径），不能报。

    ## 已知边界（写清现状，别当没有）

    1. **反斜杠紧贴字母时仍漏**：`\text{a\zzz b}` 的 m:t 是 `a\zzz\xa0b`，
       切出来的 token 是 `a\zzz` —— 不是单条命令，本判据**不命中**。抓它必须回到
       「任意位置匹配」，而那样上面三条路径公式会全被误伤。这个取舍是刻意的，
       由 `test_residue_known_blind_spot_glued` 钉住，别让人以为是 bug。
    2. **已知命令在文本模式由层 1 兜**：`\text{\eqno}` / `\text{\limits}` 这类
       （层 1 的表里有名字）到不了层 2。所以边界 1 只对**完全未知**命令成立 ——
       而未知命令本来就没有替代写法可给，报出来的价值也最低。
    3. 未展开的 `\protect\foo` 两者挤在同一 token 里也不命中；`\protect` 在层 1 表里。
    """
    out: list[str] = []
    seen: set[str] = set()
    for t in omml.iter(_m("t")):
        # str.split() 无参 = 按任意 Unicode 空白切，含 NBSP（U+00A0）——
        # latex2mathml 在 \text{} 里正是用 NBSP 填空白，这是本判据成立的前提。
        for token in (t.text or "").split():
            if token in seen:
                continue
            if _RESIDUE_RE.match(token):
                seen.add(token)
                out.append(token)
    return out


def assert_no_latex_residue(omml, latex: str, stats: dict | None = None) -> None:
    r"""层 2 的出口：产物里还有整段控制序列就抛 RuntimeError。

    调用位置在 latex_to_omml_lxml 里 split_runs_for_upright **之后**、最后一次
    assert_legal **之前**：split 会重建 m:r / m:t，扫最终树才正确。
    """
    residue = find_latex_residue(omml)
    if not residue:
        return
    _guard_hit(stats, "residue", latex, residue)
    raise LatexGuardError(
        "OMML 产物里残留了 LaTeX 控制序列（%r）：%s。\n"
        "症状：公式里印着反斜杠命令名（例如 \\zzz 整串出现），Word 里就是一串带"
        "反斜杠的怪字符，交付物上直接看得到。\n"
        "成因：这条命令 latex2mathml 不认识，被当字面量搬进了 <m:t>；它不在已知的"
        "「LaTeX 排版专属」清单里，所以只能靠这层文本扫描兜住 —— 任意未知命令都会"
        "走到这里。\n"
        "替代写法：把该命令换成等效的 LaTeX（希腊字母 / 运算符 / 结构命令），或"
        "直接从公式里删掉。要正文里的反斜杠字形请写 \\backslash。\n"
        "怎么做：按上面点名的命令逐个改源文本。若这条公式是程序拼出来的，检查拼串"
        "时有没有把排版命令一起拼进去。"
        % (latex, "、".join(residue)))


def latex_to_omml_lxml(latex: str, display: bool = False,
                       raw_spacing: bool = False, stats: dict | None = None):
    r"""LaTeX → OMML 的 lxml 元素（``m:oMath`` 或 ``m:oMathPara``）。

    返回的元素自带 xmlns:m；直接 append 进 python-docx / lxml 的 ``w:p`` 即可。

    排版修正**默认全开、不可关**（它们是排版正确性，不是可选项）：

    * **直立补全** —— ``\max\{x,y\}`` / ``\frac{dy}{dx}`` 这类被 latex2mathml 塞进
      单个 run 的文本，按语义拆段并给数字、运算符、函数名补 ``m:rPr/m:sty val="p"``；
      否则 Word 里函数名和微分算子全是斜体（见 :func:`split_runs_for_upright`）；
    * **cases 左对齐** —— ``\begin{cases}`` 转出来的 ``m:m`` 与矩阵同构，XSLT 统一给
      ``mcJc=center``；按 MathML 层的 ``columnalign`` 判出谁是 cases，只改那一个
      （见 :func:`align_cases_left`）；
    * **定界组空槽** —— ``\left(-\frac{y}{x}\right)`` 与普通括号的 ``(-\frac{y}{x})``
      这两类写法，只要组首是一元正负号，XSL 都会把那个符号判成槽分隔符，产出一个
      空的 ``m:e``（Word 画成空方块）；在 MathML 侧把「符号 + 紧随的一项」重新分组，
      XSLT 之后就只有单个非空槽（见 :func:`wrap_leading_unary_in_delim`）。

    ``stats`` 非 None 时累进这些键：``upright_split``（拆了几个 run）、
    ``upright_runs``（补了几段 ``m:sty``）、``cases_aligned``（改了几个 cases）、
    ``cases_align_skipped``（放弃原因列表）、``delim_wraps``（包了几处一元符号）。
    不传 stats 行为完全一样，只是不留痕。

    ⚠️ **本函数仍不修相量点。**

    ``latex2mathml`` 把 ``\dot{I}`` 翻成 ``m:limUpp``（上限），Word 里渲染成
    「大字 + 飘着的点」；正确形态是 ``m:acc``（U+0307 组合点）。本函数只保证
    OMML **结构合法**（``assert_legal``）、字体注入与上面两项排版修正，**不会**
    替你把相量点改对。

    手写调用方必须二选一，否则 ``\dot{I}`` 式相量一定以丑态落进交付物：

    1. 直接用 :func:`latex_to_omml_safe` —— 转换 + 修相量点 + 再自检，一次做完
       （推荐，``docx_mathify`` / ``insert_formula`` / ``docx_create`` 的 ``latex``
       block 都走这条）；
    2. 或者自己补一刀 ``fix_phasor_dot.fix_phasor_dots(el)`` —— 同目录平铺模块，
       ``from fix_phasor_dot import fix_phasor_dots`` 即可，返回改了几处。

    只有确实要「未修的原始 OMML」（例如自己做后处理）时才裸调本函数。
    """
    try:
        import latex2mathml.converter as lconv
    except ImportError as exc:  # pragma: no cover
        # 2 档要求 stderr 恰好一行（D-024），所以这条文案压成单行：
        # 原来的两行写法在新分派下会破契约。
        raise MathEnvError(
            "LaTeX→OMML 需要 latex2mathml（纯 Python、零依赖）："
            "python -m pip install latex2mathml"
        ) from exc

    from lxml import etree

    # 层 1（前置硬拒）：必须在 keep_spacing **之前**扫原始参数 —— 那一步会把
    # \quad / \, 改写成控制空格，扫改写后的串会漏判；也要在 latex2mathml 之前，
    # 这样报错能直接给替代写法，而不是让用户去读上游那个语义不明的语法异常。
    assert_no_unsupported_commands(latex, stats)

    src = latex if raw_spacing else keep_spacing(latex)
    try:
        mathml = lconv.convert(src)
    except Exception as exc:
        raise LatexSyntaxError(f"latex2mathml 转换失败: {latex!r} -> {exc}") from exc

    # 兜底：latex2mathml 偶尔会让裸 & 漏进 MathML（& 对齐符没被环境吃掉时），
    # 裸 & 会让 lxml 抛 XMLSyntaxError —— 转成 &amp; 保证文档不炸
    import re as _re
    mathml = _re.sub(r"&(?!(?:amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)",
                     "&amp;", mathml)

    # MathML 树留一份给 align_cases_left：cases 与矩阵在 OMML 侧完全同构，判据只
    # 存在于这一层，而 XSLT 会把 MathML 标记整片吃掉。实测 XSLT 是纯函数、不回改
    # 输入树，所以同一个 mml_root 转完还能继续用来判。
    mml_root = etree.fromstring(mathml.encode("utf-8"))
    if display:
        # 关键：XSLT 只在 @display="block" 时才产出 m:oMathPara
        mml_root.set("display", "block")

    # 定界组空槽：必须在 XSLT **之前**改 MathML 树。XSL 会把 \left(…\right) 组首的
    # 一元正负号当成槽分隔符（m:sepChr），产出一个空的 m:e —— Word 里是空方块，
    # 而 OMML 侧的 assert_legal 只能事后抓到它、抓不到就无法还原原意。见
    # wrap_leading_unary_in_delim 的 docstring。
    wrap_leading_unary_in_delim(mml_root, stats)

    transform, xsl_path = xslt()
    try:
        out = transform(mml_root).getroot()
    except Exception as exc:
        raise LatexSyntaxError(f"MML2OMML.XSL 应用失败（{xsl_path}）: {exc}") from exc
    if out is None:
        raise OMMLStructureError(f"XSLT 未产出 OMML: {latex!r}")

    assert_legal(out, latex)

    # 字体注入：默认 Times New Roman（可用 MATH_FONT 覆盖，设空串则不注入）
    # 注意：apply_math_font 在需要重绑命名空间时会返回新根，必须接住返回值。
    out = apply_math_font(out)
    # 顺序关键：拆 run 要 deepcopy 已注入的 w:rPr，所以必须在 apply_math_font 之后。
    split_runs_for_upright(out, stats)
    align_cases_left(mml_root, out, stats)
    # 层 2（后置残留检测）：必须在 split_runs_for_upright **之后** —— 那一步会重建
    # m:r / m:t，在它之前扫到的是中间态；扫最终树才正确。
    assert_no_latex_residue(out, latex, stats)
    # 拆 run 重写了 m:r 的结构，第一次 assert_legal 只覆盖了 XSLT 的产物，必须复检。
    assert_legal(out, latex)
    return out


def latex_to_omml_string(latex: str, display: bool = False,
                         raw_spacing: bool = False) -> str:
    """LaTeX → OMML 字符串（用于写进 document.xml 的场景）。"""
    from lxml import etree
    el = latex_to_omml_lxml(latex, display=display, raw_spacing=raw_spacing)
    return etree.tostring(el, encoding="unicode")


def assert_legal(omml, latex: str) -> None:
    r"""结构自检：抓四类会让 Word 判文档「可能已经损坏」或画错的错误。

    1. :data:`ARG_TAGS` 里内嵌 ``<m:oMath>`` —— Word 报「文件可能已经损坏」；
    2. ``<m:rad>`` 缺 ``<m:e>`` —— 根号没有内容，同样报损坏；
    3. ``<m:d>``（定界组）下出现**空的** ``<m:e>`` —— Word 不报错，而是画成一个
       空方块。这一条来自 ``MML2OMML.XSL``：组首的一元正负号被它当成槽分隔符
       （``m:sepChr``），于是前面空出一个槽。
       :func:`wrap_leading_unary_in_delim` 在 MathML 层已经修掉绝大多数形态，
       这里兜的是它修不到的两类：本来就空的组（``\left(\right)`` / ``()``
       组内 0 个项，没有可包的第二项），以及绕过 ``math_omml`` 正常链路、
       由别处手搓出来的 OMML（不会经过那次 MathML 层修复）。
    4. ``<m:oMath>`` 自身**一个子元素都没有** —— 整条公式是空的。Word 里什么都不
       显示（或画成一个空方块），而这条链路上没有任何一环会响：XML 合法、
       ``docx_validate`` 的 OPC/结构检查过、``docx_mathify`` 的 ``failed`` 只数异常。
       成因见 :func:`latex_to_omml_lxml` 上游：``latex2mathml`` 把 ``\rule`` /
       ``\hspace`` 一族转成 ``<mspace>``，而 ``MML2OMML.XSL`` 全文没有 ``mspace``
       模板，于是整条公式被吞掉；空组 ``{}`` 与「定义了却没用的 ``\def``」同样落空。
       层 1（:data:`LATEX_ONLY_COMMANDS`）已经在转换前拦掉间距/画线那一族，
       本条是它拦不到的那部分（空组、绕开正常链路的 OMML）的最后一道。

    **第 4 类与第 3 类不是一件事**，别合并：第 3 类查的是 ``m:d`` 组里的一个
    **槽**空着（公式还有别的内容，只是某个定界组缺一项）；第 4 类查的是**整条
    公式**空着（一个字都没有）。判据一个是 ``m:d`` 的直接子 ``m:e``，一个是
    ``m:oMath`` 自身 —— 前者远、后者近，混在一起就必然误伤。

    第 3 条**只**看 ``m:d`` 直接子元素里的 ``m:e``。``m:e`` 在
    ``m:num`` / ``m:den`` / ``m:fName`` 里也合法、也常为空（那是别的语义），
    一律不能判非法 —— 顺手收成「任何空 m:e」会把大量正常公式打成错误。
    ``ARG_TAGS`` 里虽然也含 ``e``，那里查的是「内嵌 m:oMath」，与本条无关。
    第 4 条同理**只**看 ``m:oMath`` 本身：``\sqrt{}`` 产的 ``<m:rad><m:e/>``、
    ``\frac{}{}`` 产的 ``<m:num/>``/``<m:den/>`` 都是「别的语义下的空」，
    ``m:oMath`` 仍有子元素（``m:rad`` / ``m:f``），不会被本条打倒 ——
    ``tests/test_math_delim_slots.py:322-341`` 正钉着这条不许误伤。
    """

    def local(el):
        t = getattr(el, "tag", "")
        return t.rsplit("}", 1)[-1] if isinstance(t, str) and t.startswith("{") else t

    for el in omml.iter():
        name = local(el)
        if name in ARG_TAGS:
            for child in el:
                if local(child) == "oMath":
                    raise OMMLStructureError(
                        f"OMML 非法：<m:{name}> 内嵌了 <m:oMath>（{latex!r}）。"
                        "这会让 Word 报「文件可能已经损坏」。"
                    )
        elif name == "rad":
            if not any(local(c) == "e" for c in el):
                raise OMMLStructureError(f"OMML 非法：<m:rad> 缺少 <m:e>（{latex!r}）。")
        elif name == "d":
            for child in el:
                if local(child) == "e" and len(child) == 0:
                    raise OMMLStructureError(
                        f"OMML 非法：<m:d> 下出现了空槽 <m:e/>（{latex!r}）。"
                        "症状：定界组里有一个空的槽，Word 把它画成一个空方块。"
                        "成因 A：组首的一元正负号被 MML2OMML.XSL 误判成槽分隔符"
                        "（m:sepChr 变成 − / + / ± / ∓），于是它前面空出一个槽 —— "
                        "正常路径上 wrap_leading_unary_in_delim 已经在 MathML 层把它"
                        "修掉了，所以还看到这一条，说明公式没走 math_omml 的正常链路"
                        "（例如别处手搓的 OMML 不会经过这层修复）。"
                        "成因 B：组里本来就是空的（() 或 \\left(\\right)），没有"
                        "内容可放 —— 工具不替它编造内容。"
                        "怎么做：先把公式补全；补全之后仍然报这个错，再检查这条公式"
                        "是不是绕过了 math_omml。"
                    )
        elif name == "oMath":
            # 第 4 类：整条公式空着。只看 m:oMath 自身**没有子元素**这一件事，
            # 绝不外扩成「任何空 m:e 都非法」（那会误伤 \sqrt{} / \frac{}{}）。
            if len(el) == 0:
                raise OMMLStructureError(
                    f"OMML 非法：这是一条空公式 —— <m:oMath/> 一个子元素都没有（{latex!r}）。"
                    "症状：Word 里这条公式什么都不显示，或画成一个空方块；文档结构"
                    "本身合法，所以 docx_validate 与转换环节都不会报错。"
                    "成因：latex2mathml 把间距 / 画线命令（\\rule、\\hspace、\\hskip、"
                    "\\kern、\\mspace、\\enspace）转成 <mspace>，而 MML2OMML.XSL "
                    "全文没有 mspace 模板，整条公式被吞掉；空组 {} 与「定义了却没用的 "
                    "\\def / \\newcommand」同样落成空公式。"
                    "怎么做：把公式补全再转；要间距就用 \\quad（1em）/ \\,（3/18em）"
                    "/ \\; / \\: / \\! / \\ （控制空格）—— 这六个是本链路唯一能"
                    "活下来的间距；要横线用段落边框或表格边框；\\def / \\newcommand "
                    "要么删掉，要么确保后面真的引用了它。"
                )


def latex_to_omml_safe(latex: str, display: bool = False,
                       raw_spacing: bool = False, stats: dict | None = None):
    r"""LaTeX → OMML，**转换 + 排版修正 + 相量点修复 + 结构自检**一次做完。

    手写调用方的默认入口。三步的顺序与 ``docx_mathify._omml_with_phasor`` 一致：

    1. :func:`latex_to_omml_lxml` —— 转成 OMML（内含 ``assert_legal``、字体注入、
       直立补全、cases 左对齐；这三项排版修正不可关，见该函数 docstring）；
    2. ``fix_phasor_dot.fix_phasor_dots`` —— ``m:limUpp`` → ``m:acc``，否则
       ``\dot{I}`` 在 Word 里是「大字 + 飘着的点」；
    3. :func:`assert_legal` —— 修完**再**自检一次。这不是重复劳动：改写本身
       也可能把结构弄坏，而 ``latex_to_omml_lxml`` 内部那次自检只覆盖转换结果。

    ``stats`` 是**同一个 dict 一路透传**给 :func:`latex_to_omml_lxml`，所以调用方
    拿到的是合并统计：``phasor_fixes``（本函数填）+ ``upright_split`` /
    ``upright_runs`` / ``cases_aligned`` / ``cases_align_skipped``（上游填）。
    修复抛异常时把 ``{"latex", "error"}`` 追加进 ``stats["phasor_errors"]``，
    并返回**未修但可用**的元素 —— 与 ``docx_mathify`` 同一套异常语义：公式本身
    是好的，相量点没修只是观感问题，不该让整篇转换失败。

    返回元素与 :func:`latex_to_omml_lxml` **同形状**：``display=False`` 是
    ``m:oMath``，``display=True`` 是 ``m:oMathPara``。
    """
    el = latex_to_omml_lxml(latex, display=display, raw_spacing=raw_spacing,
                            stats=stats)
    try:
        # 延迟 import：fix_phasor_dot 在模块顶层 import lxml，而本模块刻意让
        # lxml / latex2mathml 只在真正转换时才被拉进来（纯结构自检不需要它们）。
        from fix_phasor_dot import fix_phasor_dots
        n = fix_phasor_dots(el)
    except Exception as exc:              # 修相量点出事不该让整条公式作废
        if stats is not None:
            stats.setdefault("phasor_errors", []).append(
                {"latex": latex,
                 "error": f"{type(exc).__name__}: {str(exc)[:160]}"[:200]})
    else:
        if stats is not None:
            stats["phasor_fixes"] = stats.get("phasor_fixes", 0) + n
    assert_legal(el, latex)
    return el


if __name__ == "__main__":
    print("math_omml is a library module, import it:")
    print("    from math_omml import latex_to_omml_safe          # 推荐：转 + 修相量点")
    print("    from math_omml import latex_to_omml_lxml, latex_to_omml_string, "
          "assert_legal, apply_math_font")
    raise SystemExit(0)
