"""Force stdout/stderr to UTF-8 on Windows, without ever losing characters.

Usage — 不要再手写这一行，它已经由 `_bootstrap.py` 统一触发：

    import _bootstrap  # noqa: F401 — sys.path 自举 + minidom / UTF-8 兜底

直接 `import encoding` 也有效（会立刻执行 `_fix_stdio()`），但入口脚本走
`_bootstrap` 更省事：一个 import 同时解决 sys.path 与 stdio 两件事。
为什么必须保留这个副作用：中文文档正文（`docx_read.py` 的输出）在 cp936
流水线上会 UnicodeEncodeError 直接崩，而用户读文档正是本 skill 的主要入口。

硬化策略（P23-B 起）
--------------------
1. 先 `stream.reconfigure(encoding="utf-8", errors="backslashreplace")`；
2. **再验** `stream.encoding`（按 codec 名归一化）到底是不是 `utf-8`——
   `reconfigure` 在两类流上会"假装成功"或直接不可用：没有 `reconfigure` 的
   包装流（`AttributeError`）、把 kwargs 吃掉的空实现、以及只改了 errors
   没换编码的半实现。验不过就降级；
3. 降级 = 拿 `stream.buffer` 重绑一个
   `io.TextIOWrapper(..., encoding="utf-8", errors="backslashreplace",
   line_buffering=True)`。没有 `buffer`（如 `io.StringIO`、被替换成纯 `write`
   的桩）就放弃，保持原样——此时字符本来也到不了字节层，硬换只会崩；
4. 任何一步抛 `AttributeError` / `OSError` / `ValueError` 都只降级、不冒泡。

**绝不用 `errors="replace"`**。它把写不出去的字符静默换成 `?`，字丢在哪一行
都查不出来；`backslashreplace` 至少把 `\\ud800` 这样的转义留在输出里，可追溯。

非 Windows 的取舍
-----------------
`os.name != "nt"` 时**整段跳过**，保持原实现的行为。理由：POSIX 上 CPython 的
PEP 538/540 已经把 `C`/`POSIX` locale 强制抬到 UTF-8，真正会错的场景只剩
"终端自己是非 UTF-8 locale"——那种情况强推 UTF-8 只会把终端显示弄花，且
Linux CI 腿与各发行版默认值都已是 UTF-8，改了是净风险、零收益。

模块级只读量（诊断用，不做可变全局状态）
----------------------------------------
`STDIO_MODE`
    `"reconfigure"` / `"rewrapped"` / `"untouched"` 三选一，取 stdout 与 stderr
    里"改动最强"的那个；非 Windows 恒为 `"untouched"`。
`STDIO_MODES`
    只读映射 `{"stdout": ..., "stderr": ...}`，逐流给出模式。
"""

from __future__ import annotations

import codecs
import io
import os
import sys
from types import MappingProxyType

_TARGET_ENCODING = "utf-8"
_TARGET_ERRORS = "backslashreplace"
#: 这几个 errors 策略会**静默丢字**，就算编码已经是 UTF-8 也要重设。
_LOSSY_ERRORS = frozenset({"replace", "ignore"})

#: 见模块 docstring；import 时写入一次，之后不再变动。
STDIO_MODE = "untouched"
STDIO_MODES = MappingProxyType({"stdout": "untouched", "stderr": "untouched"})


def _encoding_name(stream: object) -> str | None:
    """返回流的编码名（codec 归一化后），取不到就返回 `None`。"""
    encoding = getattr(stream, "encoding", None)
    if not encoding:
        return None
    try:
        return codecs.lookup(encoding).name
    except (LookupError, TypeError):
        return None


def _needs_hardening(stream: object) -> bool:
    """编码不是 UTF-8，或 errors 会静默丢字 —— 两种情况都要处理。"""
    if _encoding_name(stream) != _TARGET_ENCODING:
        return True
    return getattr(stream, "errors", None) in _LOSSY_ERRORS


def _rewrap(stream: object) -> object | None:
    """用 `stream.buffer` 重绑一个 UTF-8 文本流；做不到就返回 `None`。"""
    buffer = getattr(stream, "buffer", None)
    if buffer is None:
        return None
    flush = getattr(stream, "flush", None)
    if callable(flush):
        try:
            flush()
        except (OSError, ValueError):
            return None
    try:
        return io.TextIOWrapper(
            buffer,
            encoding=_TARGET_ENCODING,
            errors=_TARGET_ERRORS,
            line_buffering=True,
        )
    except (AttributeError, OSError, ValueError):
        return None


def harden_stream(stream: object) -> tuple[object, str]:
    """把单个文本流硬化成 UTF-8，返回 `(stream, mode)`。

    `mode` 为 `"reconfigure"`（原流已就地改好）、`"rewrapped"`（换了新流，
    调用方必须把返回值写回 `sys.stdout` / `sys.stderr`）或 `"untouched"`
    （本来就好，或根本改不动——例如没有 `buffer` 的桩流）。
    """
    if stream is None or not _needs_hardening(stream):
        return stream, "untouched"
    try:
        stream.reconfigure(encoding=_TARGET_ENCODING, errors=_TARGET_ERRORS)  # type: ignore[attr-defined]
    except (AttributeError, OSError, ValueError):
        pass
    else:
        if not _needs_hardening(stream):
            return stream, "reconfigure"
    replacement = _rewrap(stream)
    if replacement is None:
        return stream, "untouched"
    return replacement, "rewrapped"


def _fix_stdio() -> None:
    global STDIO_MODE, STDIO_MODES
    if os.name != "nt":
        return
    modes: dict[str, str] = {}
    for name in ("stdout", "stderr"):
        original = getattr(sys, name, None)
        hardened, mode = harden_stream(original)
        if hardened is not original:
            setattr(sys, name, hardened)
        modes[name] = mode
    STDIO_MODES = MappingProxyType(modes)
    for candidate in ("rewrapped", "reconfigure"):
        if candidate in modes.values():
            STDIO_MODE = candidate
            break


_fix_stdio()


if __name__ == "__main__":
    print("encoding is a library module, import it:")
    print("    import encoding  # runs _fix_stdio() on import")
    print(f"    stdout={sys.stdout.encoding}  stderr={sys.stderr.encoding}")
    print(f"    STDIO_MODE={STDIO_MODE}  STDIO_MODES={dict(STDIO_MODES)}")
    raise SystemExit(0)
