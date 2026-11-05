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
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix as _encoding_fix  # noqa: F401 — UTF-8 stdio on Windows
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
                       raw_spacing: bool = False):
    """LaTeX → OMML 的 lxml 元素（``m:oMath`` 或 ``m:oMathPara``）。

    返回的元素自带 xmlns:m；直接 append 进 python-docx / lxml 的 ``w:p`` 即可。
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
