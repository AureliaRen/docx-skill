# -*- coding: utf-8 -*-
r"""P23-B2: ``tools/word_accept.py`` 的 stdio 硬化（独立副本）+ 防漂回断言的回归。

## 为什么单独测这份副本

``tools/word_accept.py`` **刻意**自带一份 stdio 硬化，不 import ``scripts/encoding.py``
（``tools/README.md`` 的理由：``scripts/`` 正在被改的时候这个探针也得能用）。
副本的代价就是**会漂**：P23-B 把 ``scripts/`` 那份改成了「验编码 + 失败降级 +
backslashreplace」，而 tools/ 这份还停在旧形态（会静默丢字的错误策略）。
本文件既锁行为，也用一条**结构断言**盯着源码别再漂回去。

## 平台口径（不是 skip）

与 ``tests/test_encoding_stdio.py`` 同一套口径：非 Windows 上硬化函数刻意不动流
（PEP 538/540 已经抬过 C locale），所以行为断言分两套——Windows 断「中文原样到字节层
且是 UTF-8」，POSIX 断「``STDIO_MODE`` 仍是 untouched、按自报编码无损往返」。
降级/就地两种分支用函数级断言（真 ``BytesIO`` + 真 ``TextIOWrapper``），两平台都跑。
"""
from __future__ import annotations

import io
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

TOOLS = Path(__file__).resolve().parents[1] / "tools"
TOOL = TOOLS / "word_accept.py"
sys.path.insert(0, str(TOOLS))

import word_accept as W  # noqa: E402

PAYLOAD = "识别 2 个片段"
IS_WINDOWS = os.name == "nt"
_ENV_KEYS = ("PYTHONIOENCODING", "PYTHONUTF8", "PYTHONLEGACYWINDOWSSTDIO")


def _env(extra=None):
    env = {k: v for k, v in os.environ.items() if k not in _ENV_KEYS}
    env.update(extra or {})
    return env


def _report(stderr: bytes) -> dict:
    text = stderr.decode("ascii", "backslashreplace")
    return dict(re.findall(r"(\w+)=(\S*)", text))


# ---------------------------------------------------------------------------
# 1. 结构断言：副本不许再漂回会丢字的策略
# ---------------------------------------------------------------------------
def test_source_never_regresses_to_a_lossy_error_policy():
    """源码里不许再出现 reconfigure 上的 ``errors="replace"`` 形态。

    这条是**防漂回**用的：副本最容易在将来被"顺手改回旧的 6 行"而没人发现，
    行为测试只在被调用时才看得见，结构断言每次跑都在看。
    """
    source = TOOL.read_text(encoding="utf-8")
    assert 'errors="replace"' not in source
    assert "errors='replace'" not in source
    assert re.search(r"reconfigure\([^)]*errors\s*=\s*[\"']replace", source) is None
    # 硬化三件套必须在场：验编码、降级重绑、留痕错误策略。
    assert "backslashreplace" in source
    assert "io.TextIOWrapper(" in source
    assert "_stdio_encoding_name" in source
    # 仍然是独立副本：不 import scripts 的 encoding。
    assert "import encoding" not in source
    assert "from encoding" not in source


def test_stdio_mode_constant_exists():
    """``STDIO_MODE`` 是诊断入口，缺了就等于又回到"猜"。"""
    assert W.STDIO_MODE in ("reconfigure", "rewrapped", "untouched")
    if not IS_WINDOWS:
        assert W.STDIO_MODE == "untouched"


# ---------------------------------------------------------------------------
# 2. 真子进程：gbk 环境下中文必须原样到字节层
# ---------------------------------------------------------------------------
CHILD = "\n".join([
    "import sys",
    f"sys.path.insert(0, {str(TOOLS)!r})",
    "import word_accept",
    f"sys.stdout.write({PAYLOAD!r})",
    "sys.stdout.flush()",
    "sys.stderr.write('ENC=%s ERR=%s MODE=%s\\n'"
    "                 % (sys.stdout.encoding, sys.stdout.errors, word_accept.STDIO_MODE))",
])


def test_child_stdout_reaches_bytes_under_gbk_env():
    """PYTHONUTF8=0 + PYTHONIOENCODING=gbk：中文原样、无 U+FFFD、错误策略不丢字。"""
    proc = subprocess.run(
        [sys.executable, "-c", CHILD],
        capture_output=True,
        env=_env({"PYTHONUTF8": "0", "PYTHONIOENCODING": "gbk"}),
        timeout=180,
    )
    assert proc.returncode == 0, proc.stderr.decode("ascii", "backslashreplace")
    info = _report(proc.stderr)
    if IS_WINDOWS:
        decoded = proc.stdout.decode("utf-8")
        assert decoded.startswith(PAYLOAD), (decoded, proc.stdout.hex())
        assert "\ufffd" not in decoded
        assert info.get("ENC") == "utf-8", info
        assert info.get("ERR") == "backslashreplace", info
        assert info.get("MODE") in ("reconfigure", "rewrapped"), info
    else:
        assert info.get("MODE") == "untouched", info
        assert proc.stdout.decode(info["ENC"]) == PAYLOAD


def test_help_runs_without_traceback():
    """``--help`` 在任何环境下都得给可读结果：exit 0、写 usage、stderr 无 traceback。

    这台机器有 Office，但 ``--help`` 不该碰 COM —— 它连 docx 都不用打开。
    """
    proc = subprocess.run(
        [sys.executable, str(TOOL), "--help"],
        capture_output=True,
        env=_env(),
        timeout=180,
    )
    assert proc.returncode == 0, proc.stderr.decode("ascii", "backslashreplace")
    out = proc.stdout.decode("utf-8")
    assert out.startswith("usage: word_accept.py"), out[:120]
    assert "--expect-pages" in out
    assert b"Traceback" not in proc.stderr


# ---------------------------------------------------------------------------
# 3. 函数级：就地 reconfigure 与降级重绑两条分支
# ---------------------------------------------------------------------------
def test_cp936_stream_is_reconfigured_in_place():
    buffer = io.BytesIO()
    stream = io.TextIOWrapper(buffer, encoding="cp936", errors="replace")

    hardened, mode = W._harden_stream(stream)

    assert mode == "reconfigure"
    assert hardened is stream
    assert hardened.encoding == "utf-8"
    assert hardened.errors == "backslashreplace"
    hardened.write(PAYLOAD + chr(0xD800))
    hardened.flush()
    raw = buffer.getvalue().decode("utf-8")
    assert raw.startswith(PAYLOAD)
    assert "\\ud800" in raw
    assert "?" not in raw


class _StubWithoutReconfigure:
    """被外部替换过的 stdout：有 buffer、没有 reconfigure、编码 cp936。"""

    encoding = "cp936"
    errors = "strict"

    def __init__(self, buffer):
        self.buffer = buffer
        self.writes = []

    def write(self, text):  # pragma: no cover - 正常路径不会走到这里
        self.writes.append(text)
        raise AssertionError("write() must not be used after rewrapping")

    def flush(self):
        pass


def test_stub_stream_without_reconfigure_is_rewrapped():
    buffer = io.BytesIO()
    stub = _StubWithoutReconfigure(buffer)

    hardened, mode = W._harden_stream(stub)

    assert mode == "rewrapped"
    assert isinstance(hardened, io.TextIOWrapper)
    assert hardened.encoding == "utf-8"
    assert hardened.errors == "backslashreplace"
    assert hardened.line_buffering is True
    hardened.write(PAYLOAD)
    hardened.flush()
    assert buffer.getvalue().decode("utf-8") == PAYLOAD
    assert stub.writes == []


def test_utf8_stream_is_left_alone_and_none_is_tolerated():
    stream = io.TextIOWrapper(io.BytesIO(), encoding="utf-8")
    assert W._harden_stream(stream) == (stream, "untouched")
    assert W._harden_stream(None) == (None, "untouched")