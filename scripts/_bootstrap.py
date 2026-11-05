#!/usr/bin/env python3
# MIT License. Part of the docx skill.
"""Import bootstrap — makes every sibling script runnable from any cwd.

Why this exists
---------------
A skill's scripts get called in wildly different ways: ``python scripts/x.py``,
``python x.py`` from inside ``scripts/``, an absolute path from a shell, or an
import from a test. Once a script does ``import helpers.foo`` or
``from docx_common import bar``, how ``sys.path`` is set decides whether it
works — and a single missing ``__init__.py`` in an intermediate layer breaks
the whole package.

The rule here is flat and boring: **all scripts live in one directory and are
imported by bare module name**. Every entry point calls ``bootstrap()`` before
importing a sibling, which guarantees its own directory is first on
``sys.path`` no matter how it was launched.

Usage at the top of an entry point::

    from _bootstrap import bootstrap
    bootstrap()

    from docx_common import iter_all_paragraphs
"""
from __future__ import annotations

import sys
from pathlib import Path

__all__ = ["bootstrap", "SCRIPT_DIR"]


SCRIPT_DIR = Path(__file__).resolve().parent


def bootstrap(verbose: bool = False) -> Path:
    """Put this script directory at the front of ``sys.path`` and return it.

    Idempotent: calling it twice is harmless.
    """
    here = str(SCRIPT_DIR)
    # Remove later duplicates so our directory genuinely wins resolution order.
    while here in sys.path[1:]:
        sys.path.remove(here)
    if not sys.path or sys.path[0] != here:
        sys.path.insert(0, here)
    if verbose:
        print(f"[_bootstrap] sys.path[0] = {sys.path[0]}", file=sys.stderr)
    return SCRIPT_DIR


# Importing this module is itself enough for most callers, so run it eagerly.
bootstrap()
