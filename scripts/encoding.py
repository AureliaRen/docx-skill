"""Reconfigure stdout/stderr to UTF-8 on Windows.  Import once per CLI script.

Usage — 不要再手写这一行，它已经由 `_bootstrap.py` 统一触发：

    import _bootstrap  # noqa: F401 — sys.path 自举 + minidom / UTF-8 兜底

直接 `import encoding` 也有效（会立刻执行 `_fix_stdio()`），但入口脚本走
`_bootstrap` 更省事：一个 import 同时解决 sys.path 与 stdio 两件事。
为什么必须保留这个副作用：中文文档正文（`docx_read.py` 的输出）在 cp936
控制台上会 UnicodeEncodeError 直接崩，而用户读文档正是本 skill 的主要入口。
"""

import os
import sys


def _fix_stdio() -> None:
    if os.name != "nt":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


_fix_stdio()


if __name__ == "__main__":
    print("encoding is a library module, import it:")
    print("    import encoding  # runs _fix_stdio() on import")
    print(f"    stdout={sys.stdout.encoding}  stderr={sys.stderr.encoding}")
    raise SystemExit(0)
