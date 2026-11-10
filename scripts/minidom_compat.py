"""`defusedxml.minidom` with a stdlib fallback for minimal Python runtimes.

`pip install defusedxml` hardens XML parsing against entity-expansion attacks, but a
bundled/embedded Python (the DeepSeek Harness runtime, a portable interpreter, a slim
container) often has no third-party packages at all, and then every script that touches
docx XML dies at import time.

Import this module *before* `import defusedxml.minidom`:

    import _bootstrap  # noqa: F401 — 平铺后的统一入口，内部会 import 本模块

（本模块原先是 `helpers/minidom_compat.py`，靠 `helpers/__init__.py` 的
`from . import minidom_compat` 触发；平铺后那条包级副作用挪进了 `_bootstrap.py`，
所以入口脚本引用的是 `_bootstrap`，不是本模块名。）

When the real package is missing it installs a stand-in `defusedxml.minidom` that wraps
`xml.dom.minidom`, so the call sites stay exactly the same. Parsing is then unhardened —
acceptable for documents this skill produced itself; `doctor()` in the writer pipeline
tells you whether the hardened parser is in use.
"""

import sys
import types

try:  # the hardened parser, when the runtime has it
    import defusedxml.minidom as minidom  # noqa: F401

    HARDENED = True
except ImportError:  # minimal runtime: fall back to the stdlib parser
    import xml.dom.minidom as minidom  # noqa: F401

    HARDENED = False
    if "defusedxml" not in sys.modules:
        _pkg = types.ModuleType("defusedxml")
        _pkg.minidom = minidom
        _pkg.__doc__ = "stdlib fallback installed by minidom_compat (flat scripts/)"
        sys.modules["defusedxml"] = _pkg
        sys.modules["defusedxml.minidom"] = minidom


if __name__ == "__main__":
    print("minidom_compat is a library module, import it:")
    print("    import minidom_compat  # or just `import _bootstrap`")
    print(f"    minidom={minidom.__name__}  HARDENED={HARDENED}")
    raise SystemExit(0)
