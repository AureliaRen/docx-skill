# -*- coding: utf-8 -*-
"""Safety regressions for ``scripts/docx_pack.py``.

Two defects are pinned here:

* **P0-C — drive-letter traversal.** ``_safe_member`` used to reject only
  ``../`` and names starting with ``/``. ``D:/x.txt`` is not POSIX-absolute, but
  ``pathlib`` on Windows treats it as rooted and ``dst / "D:/x.txt"`` discards
  ``dst``, so ``unpack()`` wrote the entry to the root of drive D. Every
  drive-letter case below therefore asserts **side effects are zero**, not just
  that an exception was raised: an implementation that writes the file first and
  then raises would pass an exception-only assertion.
* **P0-A — silent exit 0.** Running the module with arguments printed a hint and
  exited 0, doing nothing. ``--help`` must now exit 0 with real output (CI's
  ``--help`` gate requires non-empty output), and any other argument must exit 2
  on stderr.

The normal-unpack test at the end is the reverse guard: it fails if the member
check is tightened so far that ordinary archives stop extracting.
"""
from __future__ import annotations

import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx_pack import UnsafeArchiveError, pack, unpack  # noqa: E402

SCRIPT = SCRIPTS / "docx_pack.py"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _make_zip(path: Path, members: "list[tuple[str, bytes]]") -> Path:
    with zipfile.ZipFile(path, "w") as zf:
        for name, data in members:
            zf.writestr(name, data)
    return path


def _evil_archive(tmp_path: Path,
                  members: "list[tuple[str, bytes]]") -> "tuple[Path, Path, Path]":
    """Build a hostile archive in its own subdirectory.

    The archives live in ``tmp_path/in`` so the zero-side-effect assertion can
    tell "the fixture I created" from "a file unpack() leaked".
    """
    archives = tmp_path / "in"
    archives.mkdir(exist_ok=True)
    evil = _make_zip(archives / "evil.docx", members)
    return evil, tmp_path / "out", archives


def _drive_prefix(tmp_path: Path) -> str:
    """The drive that owns the checkout (``D:`` here), so the escape target is writable.

    Deliberately *not* ``tmp_path``'s drive: on this host tmp is ``C:``, whose root
    rejects writes with ``PermissionError``, so an unguarded escape would look
    "blocked" when it is only being stopped by the OS. The checkout drive is
    writable, which makes the reverse-verification honest.

    ``ntpath.splitdrive`` — the check under test — is platform-independent, so the
    literal fallback keeps the assertion meaningful off Windows.
    """
    return SCRIPTS.drive or Path(tmp_path).drive or "D:"


def _assert_no_side_effects(out_dir: Path, archives: Path) -> None:
    """Nothing may be written: not under the extraction root, not anywhere else."""
    if out_dir.exists():
        assert [p.name for p in out_dir.iterdir()] == [], (
            f"unpack() left entries behind in {out_dir}"
        )
    strays = [
        p for p in out_dir.parent.rglob("*")
        if p.is_file() and archives not in p.parents
    ]
    assert strays == [], f"unpack() wrote outside the extraction root: {strays}"


# ---------------------------------------------------------------------------
# 1. drive letter + slash — the actual P0-C escape
# ---------------------------------------------------------------------------
def test_drive_letter_with_slash_is_rejected_without_side_effect(tmp_path: Path):
    """``D:/escape.txt`` must be refused, and must not touch the drive root."""
    drive = _drive_prefix(tmp_path)
    entry = f"{drive}/p5-p0ca-escape.txt"
    escape_target = Path(f"{drive}/p5-p0ca-escape.txt")

    # Pre-condition: the escape target does not exist, otherwise the assertion
    # below would pass for the wrong reason.
    assert not escape_target.exists()

    evil, out, archives = _evil_archive(tmp_path, [(entry, b"pwned")])
    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)

    assert not escape_target.exists(), (
        f"unpack() created {escape_target} -- the drive-letter entry escaped the root"
    )
    _assert_no_side_effects(out, archives)


# ---------------------------------------------------------------------------
# 2. drive letter without a slash
# ---------------------------------------------------------------------------
def test_drive_letter_without_slash_is_rejected(tmp_path: Path):
    """``C:escape.txt`` has no separator at all and must still be refused."""
    evil, out, archives = _evil_archive(tmp_path, [("C:escape.txt", b"pwned")])
    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)
    _assert_no_side_effects(out, archives)


# ---------------------------------------------------------------------------
# 3. relative traversal — regression guard for the pre-existing behaviour
# ---------------------------------------------------------------------------
def test_relative_traversal_is_rejected(tmp_path: Path):
    evil = _make_zip(tmp_path / "evil.docx", [("../../p5-p0ca-escaped.txt", b"pwned")])
    out = tmp_path / "nested" / "out"
    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)

    assert not (tmp_path.parent.parent / "p5-p0ca-escaped.txt").exists()
    assert not any(p.is_file() for p in tmp_path.glob("**/p5-p0ca-escaped.txt"))


# ---------------------------------------------------------------------------
# 4. POSIX absolute path
# ---------------------------------------------------------------------------
def test_posix_absolute_path_is_rejected(tmp_path: Path):
    evil, out, archives = _evil_archive(tmp_path, [("/etc/p5-p0ca-escape.txt", b"pwned")])
    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)
    _assert_no_side_effects(out, archives)


# ---------------------------------------------------------------------------
# 5. reverse guard — a normal archive must still unpack
# ---------------------------------------------------------------------------
def test_normal_docx_unpacks_completely(tmp_path: Path):
    """Tightening the check must not break ordinary archives."""
    from docx import Document

    d = Document()
    d.add_paragraph("hello world")
    src = tmp_path / "normal.docx"
    d.save(src)

    with zipfile.ZipFile(src) as zf:
        expected = {i.filename for i in zf.infolist() if not i.is_dir()}
        contents = {n: zf.read(n) for n in expected}

    assert len(expected) >= 4, expected  # [Content_Types].xml, rels, word/*

    out = tmp_path / "unpacked"
    unpack(src, out)

    unpacked = {
        p.relative_to(out).as_posix()
        for p in out.rglob("*")
        if p.is_file()
    }
    assert unpacked == expected, (
        f"entry set drifted: missing={expected - unpacked}, extra={unpacked - expected}"
    )
    for name, data in contents.items():
        assert (out / name).read_bytes() == data, f"{name} content changed"
    assert "hello world" in (out / "word" / "document.xml").read_text(encoding="utf-8")


def test_pack_also_rejects_drive_letter_in_original(tmp_path: Path):
    """``pack(original=...)`` reads entry names through the same check."""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    evil = _make_zip(tmp_path / "evil.docx", [("D:/escape.txt", b"pwned")])
    with pytest.raises(UnsafeArchiveError):
        pack(src_dir, tmp_path / "out.docx", original=evil)


# ---------------------------------------------------------------------------
# 6. P0-A — no silent exit 0
# ---------------------------------------------------------------------------
def _run(*args: str) -> "subprocess.CompletedProcess[str]":
    return subprocess.run(
        [sys.executable, str(SCRIPT), *args],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        cwd=str(SCRIPTS),
    )


def test_help_exits_zero_with_output():
    r = _run("--help")
    assert r.returncode == 0, r.stderr
    assert r.stdout.strip(), "--help printed nothing; CI's gate calls that a false green"
    assert "unpack" in r.stdout and "pack" in r.stdout


def test_short_help_exits_zero_with_output():
    r = _run("-h")
    assert r.returncode == 0, r.stderr
    assert r.stdout.strip()


def test_no_arguments_exits_zero_with_output():
    r = _run()
    assert r.returncode == 0, r.stderr
    assert r.stdout.strip()


@pytest.mark.parametrize("argv", [
    ("unpack", "in.docx", "outdir"),
    ("pack", "srcdir", "out.docx"),
    ("--nonsense",),
    ("--version",),
])
def test_any_other_argument_exits_two_on_stderr(argv):
    """The old behaviour was exit 0 and nothing done -- that is the defect."""
    r = _run(*argv)
    assert r.returncode == 2, (
        f"{argv} exited {r.returncode}, expected 2 (silent success is the bug)"
    )
    assert r.stderr.strip(), f"{argv} exited 2 but said nothing on stderr"
    assert "library module" in r.stderr


def test_module_introduces_no_argparse():
    """The file must stay a hand-rolled library; CI classifies it by its shape."""
    text = SCRIPT.read_text(encoding="utf-8")
    assert "argparse" not in text
    assert 'if __name__ == "__main__":' in text