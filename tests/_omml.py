# -*- coding: utf-8 -*-
r"""OMML 可用性检测 —— **全仓库唯一一处定义**。

## 为什么单独成文件

`HANDOFF-ROUND4` §5.1 记下的教训：CI（hosted runner）既没有 Office，也没有
`MML2OMML.XSL`，`latex_to_omml_lxml` 会**静默降级成纯文字**。于是任何
「断言这段 LaTeX 变成了 `m:oMath`」的测试，在 CI 上必然失败 —— 而在装了
Word 的开发机上永远绿。这就是**假绿灯**，它真实地藏掉过一个 bug。

对策是给这类测试加跳过标记。但 `c084efd` 那次留下的是**四份各自复制**的
判定函数，且其中一份漏查 `latex2mathml`：

| 文件 | 查 `math_omml.find_xslt()` | 查 `latex2mathml` |
|---|---|---|
| `test_mixed_superscript.py` | ✅ | ❌ **漏** |
| `test_no_layout_needed.py` | ✅ | ✅ |
| `test_formula.py` | ✅ | ✅ |
| `test_units.py` | ✅ | ✅ |

后果是可预测的：在「装了 latex2mathml 但没装 Office」的机器上，三处 skip、
一处 fail —— 同一份代码给出两种结论。更重要的是，以后谁改一份忘改另一份，
跳过条件就静默分叉。

所以判定收在这里一处，四个文件 import 它（`REVIEW-CODE-R4` 应修 5）。
改判定 = 改这一个文件，四处同步生效。

## 判定口径

**两个条件都要满足**才算 OMML 可用：

1. `latex2mathml` 能 import —— LaTeX → MathML 的第一跳；
2. `math_omml.find_xslt()` 能找到 `MML2OMML.XSL` —— MathML → OMML 的第二跳。

少任一个，`latex_to_omml_lxml` 都会走降级分支。所以必须 `and`，不能只看其一。

> ⚠️ 曾有人（包括本文件的初版设计）想过「只查 XSL 就够了」—— 不对。
> `test_mixed_superscript.py` 那份漏查就是这个思路的产物，正是应修 5 要修的。
"""
from __future__ import annotations

import pytest

__all__ = [
    "omml_available",
    "needs_omml",
    "SKIP_REASON",
    "omml_unavailable_reason",
]

SKIP_REASON = (
    "needs latex2mathml + Microsoft MML2OMML.XSL (local Office install)"
)


def omml_unavailable_reason() -> str | None:
    """OMML 不可用时返回原因，可用时返回 ``None``。

    分开返回原因而不是只给 bool，是为了报告里能直接看出**缺哪一环** ——
    「CI 上为什么跳过了 87 条」这个问题，答案就藏在这里。

    `math_omml` 的 import 也包在 try 里：这个模块在 sys.path 上依赖
    `scripts/` 已入路径，四个测试文件都在 import 阶段做了
    ``sys.path.insert(0, str(SCRIPTS))``，但保险起见仍然兜住。
    """
    try:
        import latex2mathml  # noqa: F401
    except Exception:
        return "latex2mathml missing"
    try:
        import math_omml
    except Exception:
        return "math_omml missing"
    try:
        math_omml.find_xslt()
    except Exception:
        return "MML2OMML.XSL missing (no local Office install)"
    return None


def omml_available() -> bool:
    """本机能不能真的产出 OMML 公式对象。"""
    return omml_unavailable_reason() is None


def _evaluate() -> bool:
    return omml_available()


# ⚠️ 这里必须**调用** `omml_available()`，不能只传函数名。
# `pytest.mark.skipif(not omml_available, ...)` 里 `omml_available` 是函数对象、
# 恒为真值，`not` 恒为 False → 标记永不生效 → 无 Office 的机器上从 skip 变 fail。
# 这个坑本人踩过一次（1189 全绿 → 无 Office 时 87 failed），记在这里别重踩。
needs_omml = pytest.mark.skipif(
    not omml_available(), reason=SKIP_REASON)
