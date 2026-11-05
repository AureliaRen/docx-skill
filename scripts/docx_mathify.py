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

有意不转（避免造出半个公式）：纯数量（`6V`、`2Ω`）、残缺片段（`b=`、`=R`、`Uoc →`）、日期。

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
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix as _encoding_fix  # noqa: F401 — UTF-8 stdio on Windows
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

XMLSPACE = "{http://www.w3.org/XML/1998/namespace}space"

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


def _om_or_text(tex: str, fallback: str, stats: dict) -> list:
    if not tex:
        return []
    try:
        return [latex_to_omml_lxml(tex)]
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
        om = None
        try:
            om = latex_to_omml_lxml(eqnum["latex"])
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
        try:
            om = latex_to_omml_lxml(seg["latex"])
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
             "refs": 0, "bm": 9000}
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
                rows.append({"p": idx, "linear": body_tex.strip(),
                             "latex": body_tex.strip(), "context": orig[:120],
                             "num": f"({r['tag']})" if r["tag"] else ""})
            continue

        if not dry_run:
            if eqnum is not None:
                _setup_eqnum_para(p_el, mid_pos, right_pos)
            rebuild(p_el, segs, stats, eqnum=eqnum, label_map=label_map)

        if eqnum is not None:
            stats["eqnum"] += 1
        for s in maths:
            stats["math_segments"] += 1
            rows.append({
                "p": idx,
                "linear": "".join(a["ch"] for a in s["atoms"]),
                "latex": s["latex"],
                "context": orig[:120],
                "num": f"({eqnum['num']})" if eqnum is not None else "",
            })

    if not dry_run and out is not None:
        out.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(out))

    if report is not None:
        lines = ["# 公式 → LaTeX 转换报告", "",
                 f"- 源文件：`{src}`",
                 f"- 输出：`{out}`" if out else "- 输出：（dry-run，未写文件）",
                 f"- 含公式段落：{stats['paragraphs']}　转换片段：{stats['math_segments']}　"
                 f"编号公式段：{stats.get('eqnum', 0)}　对齐编号块：{stats.get('align_tables', 0)}　"
                 f"引用域：{stats.get('refs', 0)}　失败：{len(stats['failed'])}", "",
                 "| # | 段落 | 原线性写法 | 生成的 LaTeX | 编号 |", "|---|---|---|---|---|"]
        for i, r in enumerate(rows, 1):
            lines.append(f"| {i} | {r['p']} | `{r['linear'].replace('|', r'\|')}` "
                         f"| `{r['latex'].replace('|', r'\|')}` | {r.get('num', '')} |")
        if stats["script_fixes"]:
            lines += ["", "## 已自动修正的源排版错误（上下标把后面的正常字符一起包了进去）", "",
                      "| 段落 | 修正前 | 修正后 |", "|---|---|---|"]
            for f in stats["script_fixes"]:
                lines.append(f"| {f['p']} | `{f['before']}` | `{f['after']}` |")
        left = [r for r in rows if re.search(r"[_^]\{(\d+[A-Z][A-Za-z0-9]*)\}", r["latex"])]
        if left:
            lines += ["", "## ⚠️ 仍需人工确认的上下标", ""]
            for r in left:
                lines.append(f"- 段落 {r['p']}：`{r['latex']}`")
        if stats["failed"]:
            lines += ["", "## 转换失败（已保留原文）", ""]
            for f in stats["failed"]:
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
          f"引用域 {stats.get('refs', 0)}　失败 {len(stats['failed'])}　"
          f"修正上下标 {len(stats['script_fixes'])}")
    if out:
        print(f"输出: {out}")
    for f in stats["failed"][:10]:
        print("  FAIL", f["latex"], "->", f["error"][:120])
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
