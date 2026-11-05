"""Reconfigure stdout/stderr to UTF-8 on Windows. Import once per CLI script.

Python on Windows defaults stdio to the console code page (cp936 on a zh-CN
system), so printing CJK text — or a LaTeX string with non-ASCII — raises
``UnicodeEncodeError`` or silently emits mojibake. Importing this module fixes
the two streams in place.

Usage at the top of any script, right after the docstring::

    import encoding_fix  # noqa: F401 — ensure UTF-8 output on Windows
"""

import os
import sys


def fix_stdio() -> None:
    """Reconfigure stdout/stderr to UTF-8 (no-op off Windows)."""
    if os.name != "nt":
        return
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, OSError):
            pass


fix_stdio()
