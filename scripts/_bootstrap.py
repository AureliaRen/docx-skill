#!/usr/bin/env python3
# MIT License. Bootstrap for the docx skill scripts (flat layout).
"""One import that makes every sibling script reachable and installs the runtime
fallbacks the flat layout would otherwise lose.

解决什么问题
------------
这个 skill 的 ``scripts/`` **全平铺、没有子包**（新仓库硬规矩 1）。平铺带来两个后果：

1. **同目录 import 不再天然可用**。``scripts/docx_edit.py`` 里写
   ``from docx_common import ...``，只有当该脚本所在目录已经在 ``sys.path``
   上时才能解析；从别的 cwd 用绝对路径启动、或被 ``tests/`` 里的测试 import
   时就会 ``ModuleNotFoundError``。这就是 A-1 的成因（``scripts/office/`` 少一个
   ``__init__.py``，``import helpers.encoding`` 当场崩）。
2. **包级副作用没有宿主了**。原来 ``helpers/__init__.py`` 只有一行
   ``from . import minidom_compat``，作用是给没有第三方包的运行时装
   ``defusedxml.minidom`` 兜底；平铺后 ``__init__.py`` 不存在，这一行没有落点，
   无 ``defusedxml`` 的环境会静默失去兜底（硬约束 #39）。

接口
----
::

    import _bootstrap  # noqa: F401

无函数、无 CLI、无返回值；被 import 即完成工作，重复 import 幂等。

模块级只读量：

``BOOTSTRAP_DIR``
    被加进 ``sys.path`` 的目录，即 ``scripts/`` 自己。
``MINIDOM_HARDENED``
    ``True`` = 环境里有真的 ``defusedxml``（解析被加固）；
    ``False`` = 装的是 stdlib ``xml.dom.minidom`` 兜底。

踩过的坑
--------
- ``Path(__file__).parent`` 必须 ``resolve()``：带 ``..`` 的启动方式或相对路径会
  把一个解析不到的目录塞进 ``sys.path``，自举看起来成功、import 照样失败。
- ``sys.path.insert(0, ...)`` 之前必须查重，否则反复 import（测试里常见）会把它
  堆成一串重复项，长会话里能堆到几十条。
- 副作用必须在**同级其他模块被 import 之前**装好，所以下面两个 import 放在模块
  顶层并带 ``# noqa: E402``。
- 零第三方硬依赖：``defusedxml`` 缺失时走 ``minidom_compat`` 的 stdlib 兜底，
  ``encoding`` 是纯标准库；本模块在裸 CPython 上也能 import 成功。
"""
from __future__ import annotations

import sys
from pathlib import Path

_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))

BOOTSTRAP_DIR = _HERE

# 副作用 ①：给没装 defusedxml 的运行时装 xml.dom.minidom 兜底。
#           这是 helpers/__init__.py 原本的隐藏职责，平铺后挪到这里（硬约束 #39）。
import minidom_compat as _minidom_compat  # noqa: E402

MINIDOM_HARDENED = bool(getattr(_minidom_compat, "HARDENED", False))

# 副作用 ②：Windows 下把 stdout/stderr 重配成 UTF-8。
#           原来是 `import helpers.encoding` 触发的；平铺后那一行被删掉，
#           副作用由本模块统一承担——否则中文输出在 cp936 控制台上会炸。
import encoding as _encoding  # noqa: E402,F401

__all__ = ["BOOTSTRAP_DIR", "MINIDOM_HARDENED"]
