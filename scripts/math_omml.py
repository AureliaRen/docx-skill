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
from pathlib import Path

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# 内容模型：这些容器只接受「数学项」，不接受嵌套 <m:oMath>
ARG_TAGS = {"num", "den", "e", "sub", "sup", "deg", "fName", "lim"}

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
    """编译并缓存 XSLT → (transform, path)。"""
    global _XSLT
    if _XSLT is None:
        from lxml import etree
        path = find_xslt()
        _XSLT = (etree.XSLT(etree.parse(str(path))), str(path))
    return _XSLT


def keep_spacing(latex: str) -> str:
    r"""间距命令会被 XSLT 丢掉（普通空格也会），改写为 ``\ ``（控制空格 → NBSP）。"""
    import re
    return re.sub(r"\\(?:qquad|quad|,|;|:|!)", r"\\ ", latex)


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
      （见 :func:`align_cases_left`）。

    ``stats`` 非 None 时累进这些键：``upright_split``（拆了几个 run）、
    ``upright_runs``（补了几段 ``m:sty``）、``cases_aligned``（改了几个 cases）、
    ``cases_align_skipped``（放弃原因列表）。不传 stats 行为完全一样，只是不留痕。

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
        raise RuntimeError(
            "LaTeX→OMML 需要 latex2mathml（纯 Python、零依赖）：\n"
            "  python -m pip install latex2mathml"
        ) from exc

    from lxml import etree

    src = latex if raw_spacing else keep_spacing(latex)
    try:
        mathml = lconv.convert(src)
    except Exception as exc:
        raise RuntimeError(f"latex2mathml 转换失败: {latex!r} -> {exc}") from exc

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

    transform, xsl_path = xslt()
    try:
        out = transform(mml_root).getroot()
    except Exception as exc:
        raise RuntimeError(f"MML2OMML.XSL 应用失败（{xsl_path}）: {exc}") from exc
    if out is None:
        raise RuntimeError(f"XSLT 未产出 OMML: {latex!r}")

    assert_legal(out, latex)

    # 字体注入：默认 Times New Roman（可用 MATH_FONT 覆盖，设空串则不注入）
    # 注意：apply_math_font 在需要重绑命名空间时会返回新根，必须接住返回值。
    out = apply_math_font(out)
    # 顺序关键：拆 run 要 deepcopy 已注入的 w:rPr，所以必须在 apply_math_font 之后。
    split_runs_for_upright(out, stats)
    align_cases_left(mml_root, out, stats)
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
    """结构自检：抓两类会让 Word 判文档「可能已经损坏」的错误。"""

    def local(el):
        t = getattr(el, "tag", "")
        return t.rsplit("}", 1)[-1] if isinstance(t, str) and t.startswith("{") else t

    for el in omml.iter():
        name = local(el)
        if name in ARG_TAGS:
            for child in el:
                if local(child) == "oMath":
                    raise RuntimeError(
                        f"OMML 非法：<m:{name}> 内嵌了 <m:oMath>（{latex!r}）。"
                        "这会让 Word 报「文件可能已经损坏」。"
                    )
        elif name == "rad":
            if not any(local(c) == "e" for c in el):
                raise RuntimeError(f"OMML 非法：<m:rad> 缺少 <m:e>（{latex!r}）。")


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
