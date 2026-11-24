# -*- coding: utf-8 -*-
r"""P23-B: `scripts/encoding.py` 的 stdio 硬化回归（真子进程 + 真流）。

## 这份文件锁什么

汇报缺陷 5 的说法是「`_bootstrap` 里的 `reconfigure` 在父进程用非 UTF-8
打开管道时不生效」。**实测该根因不复现**：只要子进程 `import encoding`，
stdout 字节恒为 UTF-8（见 `work/p23b/probe_stdio_b.py` 的 hex 读数）。
真正会乱的是**没走 encoding 的裸脚本**——它按 cp936 写出、pwsh 按 UTF-8 读入。

所以这里不重复"reconfigure 生效"这条已证事实，只锁两件**以后能被改坏**的性质：

1. 三组环境（默认 / `PYTHONUTF8=0` / `PYTHONUTF8=0+PYTHONIOENCODING=gbk`）
   下，子进程 **stdout 字节可按 UTF-8 解码且含中文原文**，errors 策略是
   `backslashreplace` 而不是会静默丢字的 `replace`；
2. `reconfigure` 不可用的流（桩）走**降级重绑**分支：拿 `.buffer` 换一个
   UTF-8 的 `TextIOWrapper`，并且写出的中文是真的 UTF-8 字节。

## 平台口径（不是 skip，是两套不同断言）

`encoding.py` 在 `os.name != "nt"` 时**刻意保持原样**（理由见模块 docstring：
POSIX 的 PEP 538/540 已把 C locale 抬到 UTF-8，强推只会把非 UTF-8 终端弄花）。
于是本文件在 POSIX 上断言的是**另一条**同等有效的性质：`STDIO_MODE` 仍是
`"untouched"`，且子进程输出按它**自己声明的编码**无损往返——即"没被这次
改动扰动"，而不是"跳过不测"。这样 Linux CI 腿不新增 skip，也不会假绿。

## 为什么用真子进程

stdio 编码只在"进程启动时流被谁打开"这一层决定，进程内 monkeypatch
`sys.stdout` 测不到 `PYTHONUTF8` / `PYTHONIOENCODING` 对真实流的作用。
三组 env 都是真 `subprocess.run` + 真管道。
"""
from __future__ import annotations

import io
import os
import re
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

import encoding as E  # noqa: E402

PAYLOAD = "识别 2 个片段"
IS_WINDOWS = os.name == "nt"
#: 这三个变量会把子进程的 stdio 拉回非 UTF-8，必须逐个清干净再按需覆盖。
_ENV_KEYS = ("PYTHONIOENCODING", "PYTHONUTF8", "PYTHONLEGACYWINDOWSSTDIO")

ENV_CASES = [
    ("default", {}),
    ("pythonutf8-0", {"PYTHONUTF8": "0"}),
    ("pythonutf8-0-ioencoding-gbk", {"PYTHONUTF8": "0", "PYTHONIOENCODING": "gbk"}),
]


def _child_env(extra: dict[str, str] | None = None) -> dict[str, str]:
    env = {k: v for k, v in os.environ.items() if k not in _ENV_KEYS}
    env.update(extra or {})
    return env


def _run(code: str, extra: dict[str, str] | None = None) -> subprocess.CompletedProcess:
    return subprocess.run(
        [sys.executable, "-c", code],
        capture_output=True,
        env=_child_env(extra),
        timeout=180,
    )


def _report(stderr: bytes) -> dict[str, str]:
    """把子进程 stderr 里的 `K=V` 对解析成 dict（一行里可能并排多对）。"""
    text = stderr.decode("ascii", "backslashreplace")
    return dict(re.findall(r"(\w+)=(\S*)", text))


def _stdio_child(body: str) -> str:
    return (
        "import sys\n"
        f"sys.path.insert(0, {str(SCRIPTS)!r})\n"
        "import encoding\n"
        f"{body}\n"
        "sys.stdout.flush()\n"
        "sys.stderr.write('STDOUT_ENC=%s STDERR_ENC=%s STDOUT_ERR=%s MODE=%s\\n'\n"
        "                 % (sys.stdout.encoding, sys.stderr.encoding,\n"
        "                    sys.stdout.errors, encoding.STDIO_MODE))\n"
    )


@pytest.mark.parametrize("name,extra", ENV_CASES, ids=[case[0] for case in ENV_CASES])
def test_child_stdout_is_utf8_and_keeps_chinese(name, extra):
    """三组 env 下：中文原文必须原样活着到达字节层，且不留 U+FFFD。"""
    proc = _run(_stdio_child(f"sys.stdout.write({PAYLOAD!r})"), extra)
    assert proc.returncode == 0, proc.stderr.decode("ascii", "backslashreplace")
    info = _report(proc.stderr)

    if IS_WINDOWS:
        # 解码失败本身就是失败 —— 这正是"管道里乱码"的判据。
        decoded = proc.stdout.decode("utf-8")
        assert decoded.startswith(PAYLOAD), (name, decoded, proc.stdout.hex())
        assert "\ufffd" not in decoded
        assert info.get("STDOUT_ENC") == "utf-8", info
        assert info.get("STDOUT_ERR") == "backslashreplace", info
        assert info.get("MODE") in ("reconfigure", "rewrapped"), info
    else:
        # 非 Windows：刻意不动。断言"没被扰动"，按子进程自报的编码无损往返。
        assert info.get("MODE") == "untouched", info
        declared = info.get("STDOUT_ENC")
        assert declared, info
        assert proc.stdout.decode(declared) == PAYLOAD


class _StubWithoutReconfigure:
    """模拟被外面替换掉的 stdout：有 buffer、没有 reconfigure、编码是 cp936。"""

    encoding = "cp936"
    errors = "strict"

    def __init__(self, buffer):
        self.buffer = buffer
        self.writes: list[str] = []

    def write(self, text):  # pragma: no cover - 正常路径不会走到这里
        self.writes.append(text)
        raise AssertionError("write() must not be used after rewrapping")

    def flush(self):
        pass


def test_stub_stream_without_reconfigure_is_rewrapped():
    """reconfigure 抛 AttributeError → 降级用 .buffer 重绑成 UTF-8 流。"""
    buffer = io.BytesIO()
    stub = _StubWithoutReconfigure(buffer)

    hardened, mode = E.harden_stream(stub)

    assert mode == "rewrapped"
    assert hardened is not stub
    assert isinstance(hardened, io.TextIOWrapper)
    assert hardened.encoding == "utf-8"
    assert hardened.errors == "backslashreplace"
    assert hardened.line_buffering is True
    hardened.write(PAYLOAD)
    hardened.flush()
    assert buffer.getvalue().decode("utf-8") == PAYLOAD
    assert stub.writes == []


def test_lossy_errors_are_replaced_without_losing_characters():
    """cp936 + errors=replace 的流：改编码，并且写不出的字符留转义而不是 '?'。"""
    buffer = io.BytesIO()
    stream = io.TextIOWrapper(buffer, encoding="cp936", errors="replace")

    hardened, mode = E.harden_stream(stream)

    assert mode == "reconfigure"
    assert hardened is stream
    assert hardened.encoding == "utf-8"
    assert hardened.errors == "backslashreplace"
    hardened.write(PAYLOAD + chr(0xD800))  # 孤立代理项：UTF-8 编不出来
    hardened.flush()
    raw = buffer.getvalue().decode("utf-8")
    assert raw.startswith(PAYLOAD)
    assert "\\ud800" in raw  # backslashreplace 的痕迹
    assert "?" not in raw


def test_stream_already_utf8_is_left_alone():
    """已经是 utf-8 + 非丢字策略的流：一个字节都不动。"""
    stream = io.TextIOWrapper(io.BytesIO(), encoding="utf-8")
    hardened, mode = E.harden_stream(stream)
    assert mode == "untouched"
    assert hardened is stream


def test_missing_stream_is_tolerated():
    """pythonw 之类 sys.stdout 为 None 的场合：不冒泡。"""
    assert E.harden_stream(None) == (None, "untouched")


def test_fix_stdio_rebinds_when_stream_has_no_reconfigure():
    """端到端接线：桩流装上 sys.stdout 后，import encoding 真的把它换掉。"""
    child = "\n".join(
        [
            "import io, sys",
            f"sys.path.insert(0, {str(SCRIPTS)!r})",
            "buffer = io.BytesIO()",
            "stub = type('Stub', (), {'encoding': 'cp936', 'errors': 'strict',",
            "                          'buffer': buffer,",
            "                          'write': lambda self, t: None,",
            "                          'flush': lambda self: None})()",
            "sys.stdout = stub",
            "import encoding",
            "if sys.stdout is not stub:",
            "    sys.stdout.write(" + repr(PAYLOAD) + ")",
            "    sys.stdout.flush()",
            "sys.stderr.write('TYPE=%s ENC=%s MODE=%s BYTES=%s\\n'",
            "                 % (type(sys.stdout).__name__, getattr(sys.stdout, 'encoding', None),",
            "                    encoding.STDIO_MODE, buffer.getvalue().hex()))",
        ]
    )
    proc = _run(child)
    assert proc.returncode == 0, proc.stderr.decode("ascii", "backslashreplace")
    info = _report(proc.stderr)

    if IS_WINDOWS:
        assert info.get("MODE") == "rewrapped", info
        assert info.get("TYPE") == "TextIOWrapper", info
        assert info.get("ENC") == "utf-8", info
        assert bytes.fromhex(info["BYTES"]).decode("utf-8") == PAYLOAD
    else:
        assert info.get("MODE") == "untouched", info
        assert info.get("TYPE") == "Stub", info
        assert info.get("BYTES") == "", info
