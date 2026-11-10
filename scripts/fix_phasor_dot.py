# -*- coding: utf-8 -*-
r"""修 OMML 里的相量点：m:limUpp → m:acc（accent）

解决什么问题
------------
latex2mathml 把 ``\dot{X}`` 转成 limUpp（上限），点会浮在字上方且过大；Word 正确的
相量点应该是 accent（``m:acc`` + ``\u0307`` 组合点）。不跑这个补丁，所有 ``\dot{I}``
式相量在 Word 里都长成「大字 + 飘着的点」——属"不调就出丑"的必备补丁，公式链出口
必须挂上（`docx_mathify.py` / `insert_formula.py` 的调用样例见 `references/formula-guide.md`）。

接口
----
``fix_phasor_dots(omml) -> int``
    就地改写传入的 lxml 元素树（形如 ``latex_to_omml_lxml()`` 的返回值），
    返回改了几处。
``python fix_phasor_dot.py``
    自检：跑 ``\dot{I}`` / ``\dot{U}_{AB}`` / ``\dot{I}_{L}+\dot{I}_{C}``，
    打印每条改了几处与 ``acc`` / ``limUpp`` 计数（期望 limUpp=0）。

踩过的坑
--------
- 只处理「基是单个拉丁字母 + 上限是单个点」的 limUpp；真正的 ``\lim`` 上限
  （长基、多字符）一律放行，否则会把极限式改坏。
- 基里的 run 属性（``w:rPr``）必须 deepcopy 进新 ``m:acc`` 的 ``m:e``，否则字体注入
  （Times New Roman）在这一处丢字——公式里别的字符是 TNR、相量点所在的字母不是。
- 遍历前先 ``list(...)`` 固化 limUpp 列表：循环里 ``parent.replace`` 会改动正在迭代
  的树。
- 本文件原版在 ``__main__`` 里硬编码了本机 ``formula-chain`` 绝对路径来 hack
  ``sys.path``（硬约束 #38），已清；现在由顶部自举块 + ``_bootstrap`` 负责。
- docstring 必须是 raw 字符串（r 前缀）：正文里的 ``\dot`` / ``\u0307`` 在普通字符串里
  是非法转义，Python 3.12 每次 import 都会喷 SyntaxWarning。
- 但 raw docstring 的正文里**不能出现三引号字面量**（写 r + 三个双引号）——它会当场
  终结这个 docstring，报 `SyntaxError: invalid character '）'`。要提就写「r 前缀」。
"""
import re

from lxml import etree

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent))

import _bootstrap  # noqa: F401 - flat-layout bootstrap + minidom / UTF-8 fallbacks

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/officeDocument/2006/wordprocessingml/2006/main"
NS = {"m": M_NS, "w": W_NS}


def _mk(tag, **attrs):
    el = etree.Element(f"{{{M_NS}}}{tag}")
    for k, v in attrs.items():
        el.set(k, v)
    return el


def _run(text, rpr_src=None):
    """造一个 m:r，字体属性从 rpr_src 复制。"""
    r = _mk("r")
    if rpr_src is not None:
        import copy
        r.append(copy.deepcopy(rpr_src))
    t = _mk("t")
    t.text = text
    r.append(t)
    return r


def fix_phasor_dots(omml):
    """把 limUpp 形式的上方小圆点改成 accent 形式。就地修改并返回。"""
    n = 0
    for lim in list(omml.iter(f"{{{M_NS}}}limUpp")):
        parent = lim.getparent()
        if parent is None:
            continue
        base_el = lim.find(f"{{{M_NS}}}e")
        lim_el = lim.find(f"{{{M_NS}}}lim")
        if base_el is None:
            continue

        base_text = "".join(base_el.itertext())
        # 只处理基是单个拉丁字母、且上限是单个点的情形
        if not (len(base_text) == 1 and base_text.isalpha()):
            continue
        lim_text = "".join(lim_el.itertext()) if lim_el is not None else ""
        if lim_text not in ("\u0307", "\u02d9", "\u0309", "\u02da", "."):
            continue

        # 取基的 run 属性
        rpr_src = None
        br = base_el.find(f"{{{M_NS}}}r")
        if br is not None:
            rpr_src = br.find(f"{{{W_NS}}}rPr")

        acc = _mk("acc")
        accpr = _mk("accPr")
        chr_el = _mk("chr", **{f"{{{M_NS}}}val": "\u0307"})
        accpr.append(chr_el)
        acc.append(accpr)

        acc_e = _mk("e")
        for child in list(base_el):
            acc_e.append(child)
        acc.append(acc_e)

        parent.replace(lim, acc)
        n += 1
    return n


if __name__ == "__main__":
    # 自检：\dot{} 的回归用例。原版在这里硬编码了本机 formula-chain 绝对路径
    # （硬约束 #38），已清掉——同目录自举由文件顶部的 sys.path.insert 负责。
    from math_omml import latex_to_omml_lxml

    for t in [r"\dot{I}", r"\dot{U}_{AB}", r"\dot{I}_{L}+\dot{I}_{C}"]:
        el = latex_to_omml_lxml(t)
        n = fix_phasor_dots(el)
        xml = etree.tostring(el).decode()
        print(f"{t!r:28} 修复 {n} 处  acc={xml.count('<m:acc')}  limUpp={xml.count('<m:limUpp')}")
