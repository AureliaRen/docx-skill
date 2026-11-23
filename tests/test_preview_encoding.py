# -*- coding: utf-8 -*-
r"""P19-A: ``docx_preview.convert_to_pdf`` must never raise ``UnicodeDecodeError``.

## The defect this file pins

Before the fix, ``convert_to_pdf`` called::

    subprocess.run(cmd, capture_output=True, text=True,
                   timeout=timeout, check=False)

``text=True`` without ``encoding`` decodes the child's pipes with
``locale.getpreferredencoding(False)`` -- **cp936** on a Chinese Windows.
``soffice`` echoes the path it is converting, so a Chinese directory name in
the input path lands in the pipe as UTF-8 and the decode dies **inside
subprocess' own reader thread**::

    Exception in thread Thread-1 (_readerthread):
      File "...\subprocess.py", line 1601, in _readerthread
        buffer.append(fh.read())
    UnicodeDecodeError: 'gbk' codec can't decode byte 0xaf in position 46:
    illegal multibyte sequence

Two things make that the worst possible shape, and both are asserted below:

1. the traceback goes to stderr, so a caller reads it as a failure while the
   process **still exits 0**;
2. the exception kills the reader thread, so ``_tail(proc)`` sees an **empty**
   pipe and the real error sentence is lost.

## Why these tests do not need LibreOffice

The bug lives in *how the pipes are decoded*, not in what soffice says. So the
tests drive the real ``convert_to_pdf`` through the real
``subprocess.run`` machinery and only swap the **executable**: a ``sys.executable
-c`` one-liner stands in for soffice and writes the hostile bytes to stdout.
``subprocess``'s reader threads, the locale-driven decode, and the
``UnicodeDecodeError`` are all genuinely exercised -- nothing is mocked away.

The payload is ``预览编码``. Its UTF-8 bytes are ``e9 a2 84 e8 a7 88 e7 bc 96
e7 a0 81``; ``0x81`` is unassigned in cp1252 and the sequence is illegal in
cp936, so the same bytes fail **ascii, cp936 and cp1252** alike. That is what
makes these cases red on a UTF-8 CI runner too, rather than green by accident
because the runner's locale happened to be UTF-8 (the classic false green that
``tests/_omml.py`` documents a sibling of).

Only the last group is skippable, and it gates on **both** LibreOffice and
PyMuPDF through the module's own probes -- same discipline as
``test_docx_preview.py:439``. A CI runner has neither, so those skip honestly
instead of failing.
"""
from __future__ import annotations

import ast
import json
import os
import subprocess
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS))

import docx_preview as dp  # noqa: E402

SCRIPT = SCRIPTS / "docx_preview.py"
MODULE_SRC = SCRIPT.read_text(encoding="utf-8")

# UTF-8 bytes valid in neither cp936 nor cp1252 nor ascii -- see module docstring.
HOSTILE_TEXT = "预览编码"
HOSTILE_BYTES = HOSTILE_TEXT.encode("utf-8")


# ------------------------------------------- 0. the textual fact we are pinning
def test_convert_to_pdf_passes_encoding_and_errors():
    """The call site must carry ``encoding="utf-8", errors="replace"``.

    Read through ``ast`` rather than by substring so a commented-out or
    docstring-only mention cannot satisfy it.  This is the guard that runs on
    **every** host: the behavioural cases below need a non-UTF-8 locale, this
    one does not.
    """
    tree = ast.parse(MODULE_SRC)
    calls = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        owner = getattr(func, "value", None)
        if (isinstance(func, ast.Attribute) and func.attr == "run"
                and isinstance(owner, ast.Name) and owner.id == "subprocess"):
            calls.append(node)
    assert calls, "docx_preview must still spawn soffice through subprocess.run"

    target = None
    for call in calls:
        kwargs = {kw.arg for kw in call.keywords if kw.arg}
        if {"capture_output", "text"} & kwargs:
            target = call
            break
    assert target is not None, "the soffice call lost capture_output/text"

    kwargs = {kw.arg: kw.value for kw in target.keywords if kw.arg}
    assert "encoding" in kwargs, (
        "soffice output is decoded with the locale -- the cp936 crash is back")
    assert isinstance(kwargs["encoding"], ast.Constant)
    assert kwargs["encoding"].value.lower().replace("-", "") == "utf8"
    assert "errors" in kwargs, "no errors= fallback: a weird byte raises again"
    assert isinstance(kwargs["errors"], ast.Constant)
    assert kwargs["errors"].value == "replace"


# --------------------------------------------------------------------- helpers
def foreign_codec_available(codec: str) -> bool:
    """Can this interpreter even decode with *codec*?

    ``cp936`` is absent from many Linux Python builds (it raises
    ``LookupError``, not ``UnicodeDecodeError``). Those cases must skip rather
    than fail, because the *bug* they pin is Windows-locale-specific.
    """
    try:
        "".encode(codec)
    except LookupError:
        return False
    return True


def fake_soffice(tmp_path: Path, body: str) -> str:
    """A stand-in executable that runs *body* under this very interpreter.

    ``convert_to_pdf`` passes the document path as ``argv`` and takes the
    binary as a single string, so the stand-in has to be **spawnable as a file**
    -- pointing it at a bare ``.py`` fails with ``WinError 193`` on Windows
    (the shell association is not consulted by ``CreateProcess``).  So the
    payload goes in a real script and a two-line launcher wraps it:

    * Windows: a ``.bat`` -- ``CreateProcess`` honours ``.bat``;
    * POSIX: a ``#!/bin/sh`` script with the exec bit, which is how a CI
      runner reaches ``sys.executable`` without assuming it is on ``PATH``.
    """
    impl = tmp_path / "fake_soffice_impl.py"
    impl.write_text("# -*- coding: utf-8 -*-\n" + body, encoding="utf-8",
                    newline="\n")

    if os.name == "nt":
        launcher = tmp_path / "fake_soffice.bat"
        launcher.write_bytes(
            ('@echo off\r\n"%s" "%s" %%*\r\n'
             % (sys.executable, impl)).encode("utf-8"))
    else:
        launcher = tmp_path / "fake_soffice.sh"
        launcher.write_text(
            '#!/bin/sh\nexec "%s" "%s" "$@"\n' % (sys.executable, impl),
            encoding="utf-8", newline="\n")
        launcher.chmod(0o755)
    return str(launcher)


def run_convert(docx_path: Path, out_dir: Path, body: str,
                timeout: int = 60) -> tuple[str, float]:
    """Call ``convert_to_pdf`` with the fake child standing in for soffice."""
    fake = fake_soffice(out_dir, body)
    return dp.convert_to_pdf(str(docx_path), str(out_dir), fake, timeout=timeout)


def make_docx(path: Path) -> Path:
    """A real .docx -- python-docx is a hard dependency of the skill."""
    from docx import Document

    doc = Document()
    doc.add_paragraph("编码回归样本。")
    doc.save(str(path))
    return path


def child_script(pdf_name: str = "in.pdf", stdout_bytes: bytes = b"",
                 stderr_bytes: bytes = b"", then: str = "") -> str:
    """Source for the stand-in soffice: write a PDF, then emit raw bytes.

    Built by concatenation, never ``%``-formatting: the payload contains
    ``%PDF`` and a literal format character would blow up in the test itself
    rather than in the code under test.
    """
    lines = ["import pathlib, sys",
             "out = pathlib.Path(sys.argv[sys.argv.index('--outdir') + 1])"]
    if pdf_name:
        lines.append("(out / " + repr(pdf_name)
                     + ").write_bytes(b'%PDF-1.7\\n')")
    if stdout_bytes:
        lines.append("sys.stdout.buffer.write(" + repr(stdout_bytes) + ")")
        lines.append("sys.stdout.flush()")
    if stderr_bytes:
        lines.append("sys.stderr.buffer.write(" + repr(stderr_bytes) + ")")
        lines.append("sys.stderr.flush()")
    if then:
        lines.append(then)
    return "\n".join(lines) + "\n"


# ------------------------------------- driving the bug inside a localised child
# Why a child process and not ``monkeypatch``: the codec used for a pipe is
# resolved by ``TextIOWrapper``'s C implementation from the *locale*, and no
# Python-level patch reaches it.  ``io.text_encoding`` patched, locale patched,
# ``PYTHONUTF8=1`` inherited -- all still decode as UTF-8 and the bug hides.
# Measured: ``monkeypatch.setattr("locale.getpreferredencoding", ...)`` has no
# effect at all, so a same-process test would be a false green on every runner.
# What *does* work is spawning this file itself with UTF-8 mode off, so the
# child's ``locale.getpreferredencoding(False)`` is genuinely cp936.

DRIVER = r'''
import json, sys, threading
sys.path.insert(0, r"{scripts}")
import locale
import docx_preview as dp

# The defect does not propagate out of ``subprocess.run``: the reader thread
# dies, ``communicate()`` returns the empty buffer, and the caller sees an
# ordinary result.  So a plain try/except around the call CANNOT see it -- the
# only reliable witness is the thread's own excepthook.
thread_errors = []
threading.excepthook = lambda args: thread_errors.append(
    "%s: %s" % (getattr(args.exc_type, "__name__", "?"), args.exc_value))

result = {{"preferred": locale.getpreferredencoding(False)}}
try:
    pdf, seconds = dp.convert_to_pdf(r"{docx}", r"{out}", r"{fake}", timeout={timeout})
    result["status"] = "ok"
    result["pdf"] = pdf
except dp.PreviewError as exc:
    result["status"] = "PreviewError"
    result["error"] = " ".join(str(exc).split())
except Exception as exc:
    result["status"] = type(exc).__name__
    result["error"] = " ".join(str(exc).split())
result["thread_errors"] = thread_errors
sys.stdout.buffer.write(json.dumps(result, ensure_ascii=True).encode("ascii"))
'''

LOCALE_CACHE: dict = {}


def localised_child_preferred_encoding() -> str:
    """The encoding a child gets with ``PYTHONUTF8=0`` on this machine."""
    if "enc" in LOCALE_CACHE:
        return LOCALE_CACHE["enc"]
    env = dict(os.environ)
    env["PYTHONUTF8"] = "0"
    env.pop("PYTHONIOENCODING", None)
    probe = subprocess.run(
        [sys.executable, "-X", "utf8=0", "-c",
         "import locale;print(locale.getpreferredencoding(False))"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        env=env)
    LOCALE_CACHE["enc"] = probe.stdout.strip() or "cp936"
    return LOCALE_CACHE["enc"]


def locale_cannot_show_the_bug() -> str | None:
    """Reason string when this host cannot reproduce the defect, else ``None``.

    The defect is *only* reachable when the process locale is a non-UTF-8
    codec.  A runner with ``LANG=C.UTF-8`` decodes the payload happily, so the
    bug is unobservable there -- an honest skip, not a broken test.  The
    ``ast`` test above is the locale-independent guard that still runs.
    """
    enc = localised_child_preferred_encoding().lower().replace("-", "")
    if enc in ("utf8", "utf-8"):
        return "host forces UTF-8 decoding; the locale bug is unreachable"
    if not foreign_codec_available(localised_child_preferred_encoding()):
        return "codec %r unavailable" % enc
    try:
        HOSTILE_BYTES.decode(localised_child_preferred_encoding())
    except UnicodeDecodeError:
        return None
    return "payload decodes as %r; cannot distinguish fix from bug" % enc


LOCALE_SKIP_REASON = locale_cannot_show_the_bug()
needs_locale = pytest.mark.skipif(LOCALE_SKIP_REASON is not None,
                                  reason=LOCALE_SKIP_REASON or "")


def run_in_localised_child(tmp_path: Path, body: str,
                           timeout: int = 60) -> dict:
    """Run ``convert_to_pdf`` in a child whose locale encoding is not UTF-8.

    Returns the driver's JSON, including ``thread_errors`` -- the list of
    exceptions that reached ``threading.excepthook``, which is where the
    original ``UnicodeDecodeError`` actually landed.
    """
    out = tmp_path / "out"
    out.mkdir(exist_ok=True)
    fake = fake_soffice(out, body)
    docx = make_docx(tmp_path / "in.docx")
    driver = tmp_path / "driver.py"
    driver.write_text(
        DRIVER.format(scripts=str(SCRIPTS),
                      fake=fake.replace("\\", "\\\\"),
                      docx=str(docx).replace("\\", "\\\\"),
                      out=str(out).replace("\\", "\\\\"),
                      timeout=timeout),
        encoding="utf-8", newline="\n")
    env = dict(os.environ)
    env["PYTHONUTF8"] = "0"
    env.pop("PYTHONIOENCODING", None)
    proc = subprocess.run(
        [sys.executable, "-X", "utf8=0", str(driver)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        env=env)
    assert proc.returncode == 0, proc.stderr
    assert "UnicodeDecodeError" not in proc.stderr, proc.stderr
    return json.loads(proc.stdout)


def assert_no_thread_exception(got: dict) -> None:
    """The core judgement: no reader thread died, whatever the bytes were."""
    assert got["thread_errors"] == [], (
        "a reader thread raised: %s" % (got["thread_errors"],))


# ------------------------------- 1. the happy path still works (no regression)
def test_clean_output_is_decoded_and_pdf_returned(tmp_path):
    """UTF-8 output is read normally and the produced PDF is returned."""
    docx = make_docx(tmp_path / "in.docx")
    out = tmp_path / "out"
    out.mkdir()
    body = child_script(stdout_bytes="convert 中文路径 as a Writer document\n"
                                    .encode("utf-8"))
    pdf, seconds = run_convert(docx, out, body)
    assert Path(pdf).is_file()
    assert Path(pdf).read_bytes().startswith(b"%PDF-")
    assert seconds >= 0


# --------------------------- 2. the regression: non-decodable output must not raise
@needs_locale
def test_undecodable_stdout_does_not_raise(tmp_path):
    """The reported crash: a UTF-8 path echo read back with the locale codec.

    Pre-fix this ends with a ``UnicodeDecodeError`` in the reader thread (and
    in ``got["thread_errors"]``); post-fix the conversion simply succeeds.
    """
    got = run_in_localised_child(tmp_path, child_script(stdout_bytes=HOSTILE_BYTES))
    assert_no_thread_exception(got)
    assert got["status"] == "ok", got
    assert Path(got["pdf"]).is_file()


@needs_locale
def test_failure_message_survives_hostile_bytes(tmp_path):
    """The second half of the defect: ``_tail`` must still see the error text.

    Pre-fix the reader thread died, the pipe came back empty, and
    ``PreviewError`` degraded to ``"(no output)"`` -- the sentence from soffice
    was silently discarded.
    """
    body = child_script(pdf_name="",
                        stdout_bytes=HOSTILE_BYTES + b"Error: conversion died")
    got = run_in_localised_child(tmp_path, body)
    assert_no_thread_exception(got)
    assert got["status"] == "PreviewError", got
    assert "wrote no PDF" in got["error"]
    assert "Error: conversion died" in got["error"], (
        "the child's error text was lost -- the reader thread died again")
    assert "(no output)" not in got["error"]


@needs_locale
def test_hostile_stderr_does_not_raise(tmp_path):
    """stderr is a separate pipe with its own reader thread; pin it too."""
    got = run_in_localised_child(tmp_path, child_script(stderr_bytes=HOSTILE_BYTES))
    assert_no_thread_exception(got)
    assert got["status"] == "ok", got
    assert Path(got["pdf"]).is_file()


def test_timeout_still_wins_over_encoding(tmp_path):
    """``TimeoutExpired`` must stay a ``PreviewError``, not a decode surprise."""
    docx = make_docx(tmp_path / "in.docx")
    out = tmp_path / "out"
    out.mkdir()
    body = child_script(stdout_bytes=HOSTILE_BYTES,
                        then="import time; time.sleep(30)")
    with pytest.raises(dp.PreviewError) as excinfo:
        run_convert(docx, out, body, timeout=1)
    assert "did not finish" in str(excinfo.value)


def test_oserror_still_wins(tmp_path):
    """A non-spawnable soffice keeps its own ``PreviewError`` wording."""
    docx = make_docx(tmp_path / "in.docx")
    out = tmp_path / "out"
    out.mkdir()
    missing = str(tmp_path / "no-such-soffice.exe")
    with pytest.raises(dp.PreviewError) as excinfo:
        dp.convert_to_pdf(str(docx), str(out), missing, timeout=10)
    assert "cannot run soffice" in str(excinfo.value)


# --------------------------------------------------- 3. real soffice (skippable)
needs_render = pytest.mark.skipif(
    dp.py_mupdf_reason() is not None or dp.find_soffice()[0] is None,
    reason="needs LibreOffice + PyMuPDF (absent on the CI runners)")


@needs_render
def test_real_soffice_with_chinese_path_under_cp936(tmp_path, monkeypatch):
    """End-to-end with the genuine binary and the locale that broke it.

    Directory *and* file name carry Chinese characters, so soffice echoes UTF-8
    bytes into a pipe that the pre-fix code would decode as cp936.
    """
    monkeypatch.setattr("locale.getpreferredencoding", lambda *a, **k: "cp936")
    folder = tmp_path / "中文目录"
    folder.mkdir()
    docx = make_docx(folder / "阅读批注作业.docx")
    soffice, why = dp.find_soffice()
    assert soffice is not None, why
    out = tmp_path / "out"
    pdf, _ = dp.convert_to_pdf(str(docx), str(out), soffice, timeout=180)
    assert Path(pdf).is_file()
    assert Path(pdf).read_bytes().startswith(b"%PDF-")