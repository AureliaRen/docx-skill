# -*- coding: utf-8 -*-
"""Safety regressions for ``scripts/docx_pack.py``.

Five defect families are pinned here:

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
* **P6-FIX-C5C — device names.** ``NUL`` (and ``nul``, ``CON.txt``, ``word/CON``)
  is resolved by Win32 to a character device, so the write was swallowed and
  ``unpack()`` still reported success; ``CON`` materialised a 1-byte file. Both
  are now refused.
* **P6-FIX-C5C — trailing dot or space.** Win32 strips it, so ``"evil.txt "``
  landed as ``"evil.txt"`` and two entries could collide on one file.
* **P6-FIX-C5C — silent ``original=`` downgrade.** ``pack(original=<missing>)``,
  ``original=""`` and ``original=<directory>`` used to be skipped without a
  word, producing a .docx that had lost every part the caller did not touch.

The normal-unpack test at the end of the P0 section is the reverse guard: it
fails if the member check is tightened so far that ordinary archives stop
extracting.
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


# ---------------------------------------------------------------------------
# 7. P6-FIX-C5C — names Windows rewrites behind our back
# ---------------------------------------------------------------------------
# Everything below needs an entry name that Python's own ``zipfile`` cannot
# produce on Windows: ``ZipInfo.__init__`` runs every name through
# ``_sanitize_filename``, which rewrites "\\" to "/" when ``os.sep == "\\"``.
# ``writestr("..\\evil.txt")`` therefore stores "../evil.txt", and the backslash
# and device-name shapes below would silently degrade into the cases that were
# already covered. So these fixtures are hand-written ZIPs: local header, stored
# payload, central directory, EOCD — the name goes in as raw UTF-8 bytes.
def _raw_zip(path: Path, name: str, data: bytes = b"pwned") -> Path:
    """Write a one-entry archive whose entry name is stored verbatim."""
    import struct
    import zlib

    nb = name.encode("utf-8")
    crc = zlib.crc32(data) & 0xFFFFFFFF
    time_, date_ = 0, ((2020 - 1980) << 9) | (1 << 5) | 1
    local = struct.pack("<IHHHHHIIIHH", 0x04034B50, 20, 0, 0, time_, date_,
                        crc, len(data), len(data), len(nb), 0) + nb
    central = struct.pack("<IHHHHHHIIIHHHHHII", 0x02014B50, 20, 20, 0, 0,
                          time_, date_, crc, len(data), len(data), len(nb),
                          0, 0, 0, 0, 0x20, 0) + nb
    eocd = struct.pack("<IHHHHIIH", 0x06054B50, 0, 0, 1, 1,
                       len(central), len(local) + len(data), 0)
    path.write_bytes(local + data + central + eocd)
    return path


def _landed(root: Path) -> "list[str]":
    """Real directory entries under ``root``.

    ``Path.exists()`` lies about device names -- stat on ``out/NUL`` succeeds
    because the device exists, not because unpack() wrote anything -- so landing
    evidence has to come from a directory walk.
    """
    if not root.exists():
        return []
    return sorted(str(p.relative_to(root)) for p in root.rglob("*") if p.is_file())


@pytest.mark.parametrize("entry", [
    "CON", "NUL", "AUX", "PRN", "COM1", "LPT1", "CONIN$", "CONOUT$",
    "word/CON", "CON.txt", "nul",
])
def test_device_name_entries_are_rejected(tmp_path: Path, entry: str):
    """A device name is not a part: NUL swallowed the write and unpack() still
    reported success, so the part vanished from an "unpacked" document."""
    archives = tmp_path / "in"
    archives.mkdir()
    evil = _raw_zip(archives / "evil.docx", entry, b"x")
    out = tmp_path / "out"

    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)

    assert _landed(out) == [], f"{entry!r} left {_landed(out)} behind"


@pytest.mark.parametrize("entry", ["evil.txt ", "evil2.txt.", "evil3. ."])
def test_trailing_dot_or_space_is_rejected(tmp_path: Path, entry: str):
    """Win32 strips a trailing dot/space, so "evil.txt " lands as "evil.txt" and
    two distinct entries collide on one file."""
    archives = tmp_path / "in"
    archives.mkdir()
    evil = _raw_zip(archives / "evil.docx", entry)
    out = tmp_path / "out"

    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)

    assert _landed(out) == []


@pytest.mark.parametrize("entry", [
    "..\\evil.txt",
    "a\\..\\..\\evil.txt",
    "\\\\server\\share\\evil.txt",
])
def test_backslash_entries_never_land_outside_root(tmp_path: Path, entry: str):
    """The backslash branch is a POSIX-only guard, but the promise holds on both
    platforms: on Windows ``_sanitize_filename`` turns the name into the
    forward-slash form and the normpath/drive check picks it up instead.
    Whichever branch fires, nothing may be written."""
    archives = tmp_path / "in"
    archives.mkdir()
    evil = _raw_zip(archives / "evil.docx", entry)
    out = tmp_path / "out"

    with pytest.raises(UnsafeArchiveError):
        unpack(evil, out)

    assert _landed(out) == []
    strays = [p for p in tmp_path.rglob("evil.txt") if archives not in p.parents]
    assert strays == [], f"unpack() wrote {entry!r} outside the extraction root: {strays}"


# ---------------------------------------------------------------------------
# 8. P6-FIX-C5C — pack(original=...) must not degrade silently
# ---------------------------------------------------------------------------
def test_pack_original_missing_raises(tmp_path: Path):
    """A missing original used to be skipped: the .docx came out looking fine
    with every part the caller did not touch already gone."""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    out = tmp_path / "out.docx"

    with pytest.raises(FileNotFoundError):
        pack(src_dir, out, original=tmp_path / "missing.docx")
    assert not out.exists(), "pack() wrote an archive even though original= was bogus"


def test_pack_original_empty_string_raises(tmp_path: Path):
    """``""`` is the classic argparse default; it must not read as "no original"."""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    out = tmp_path / "out.docx"

    with pytest.raises(ValueError):
        pack(src_dir, out, original="")
    assert not out.exists()


def test_pack_original_directory_raises(tmp_path: Path):
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    adir = tmp_path / "adir"
    adir.mkdir()

    with pytest.raises(FileNotFoundError):
        pack(src_dir, tmp_path / "out.docx", original=adir)


def test_pack_original_not_a_zip_raises_with_path(tmp_path: Path):
    """Pre-existing behaviour (BadZipFile) is kept, but the message now says
    which file was rejected."""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    notzip = tmp_path / "notzip.docx"
    notzip.write_text("i am not a zip", encoding="utf-8")

    with pytest.raises(zipfile.BadZipFile) as excinfo:
        pack(src_dir, tmp_path / "out.docx", original=notzip)
    assert "notzip.docx" in str(excinfo.value)


def test_pack_without_original_still_packs(tmp_path: Path):
    """Reverse guard: opting out has to stay a supported, silent-free path."""
    src_dir = tmp_path / "src"
    src_dir.mkdir()
    (src_dir / "a.txt").write_text("a", encoding="utf-8")
    out = tmp_path / "out.docx"

    pack(src_dir, out)
    assert zipfile.ZipFile(out).namelist() == ["a.txt"]