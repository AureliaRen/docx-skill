"""Minimal OPC (zip) pack/unpack helpers for .docx editing.

A .docx is a zip archive of XML parts. Editing one usually means: unzip, change
a part, rezip. That is all this module does — deliberately.

Design notes
------------
* **Byte-faithful rezip.** Parts are written back in their original order, and
  entries that were already compressed with the STORED method stay stored.
  Rewriting every entry with a fresh deflate level changes the file size and can
  make byte-level diffs noisy, which matters when a document is under version
  control.
* **Path-traversal refusal.** Zip entries whose names escape the extraction root
  (`../`, absolute paths, backslashes) are rejected rather than written. An
  archive is untrusted input; a crafted .docx must not be able to write outside
  the directory it is unpacked into.
  "Absolute" here includes **Windows drive-letter forms** — `D:/x.txt`, `D:x.txt`
  and `//server/share/x.txt`. They are not POSIX-absolute, but `pathlib` on
  Windows treats them as rooted, and `dst / "D:/x.txt"` silently *discards* the
  left-hand side, so the entry would land outside the extraction root. Drive
  prefixes are detected with `ntpath.splitdrive`, which is a no-op on POSIX (it
  returns an empty drive, so Linux behaves exactly as before).
* **Windows name rules are part of the contract.** Two more entry shapes are
  refused because the NTFS/Win32 layer rewrites the name behind our back, and a
  name that does not survive the round trip breaks `unpack`→`pack` on documents
  that are supposed to be byte-faithful:
  * **Reserved device names** — `CON`, `NUL`, `AUX`, `PRN`, `COM0`-`COM9`,
    `LPT0`-`LPT9`, `CONIN$`, `CONOUT$`, tested case-insensitively on the last
    component with any extension ignored (`nul`, `CON.txt`). `NUL` accepts the
    write and discards it, so `unpack()` used to report success for a part that
    never reached the disk.
  * **Trailing dot or space** in the last component — `evil.txt ` and `evil3. .`
    are created as `evil.txt` and `evil3`, so two distinct entries can collide on
    one file and the directory tree stops matching the archive.
* **No dependency on `docx_utils`.** Kept standalone so it can be used by any
  script without dragging in the XML helper surface.

Usage::

    from docx_pack import unpack, pack
    unpack("in.docx", "unpacked/")
    ...                     # edit files under unpacked/
    pack("unpacked/", "out.docx", original="in.docx")
"""
from __future__ import annotations

import ntpath
import posixpath
import shutil
import sys
import zipfile
from pathlib import Path

__all__ = ["unpack", "pack", "UnsafeArchiveError"]


class UnsafeArchiveError(ValueError):
    """Raised when an archive entry would escape the extraction root."""


#: Names the Win32 layer resolves to a character device no matter which
#: directory they appear in. `COM0`/`LPT0` are reserved on current Windows too.
_RESERVED_DEVICE_NAMES = frozenset(
    ["CON", "PRN", "AUX", "NUL", "CONIN$", "CONOUT$"]
    + [f"COM{i}" for i in range(10)]
    + [f"LPT{i}" for i in range(10)]
)


def _safe_member(name: str) -> str:
    """Normalize a zip entry name, rejecting anything outside the root."""
    if not name:
        raise UnsafeArchiveError("empty archive entry name")
    # Drive-prefix check must come before the POSIX ones: "D:/x.txt" and "D:x.txt"
    # are not POSIX-absolute, yet pathlib on Windows treats both as rooted and
    # `dst / "D:/x.txt"` drops `dst` entirely. splitdrive is empty on POSIX, so
    # this branch never fires on Linux. It also reports UNC shares ("//host/share")
    # with a non-empty drive, hence "prefix" rather than "letter".
    drive, _ = ntpath.splitdrive(name)
    if drive:
        raise UnsafeArchiveError(
            f"archive entry is absolute (drive prefix): {name!r}"
        )
    # Not dead code, but conditionally live -- keep it, see the note below.
    # `zipfile.ZipInfo.__init__` runs every name through `_sanitize_filename`,
    # which on Windows (``os.sep == "\\"``) rewrites "\\" to "/" *before*
    # `infolist()` ever returns it. So on Windows this branch cannot fire: a
    # crafted "a\\..\\..\\evil.txt" reaches us as "a/../../evil.txt" and is
    # caught by the normpath check instead. On POSIX ``os.sep`` is "/", nothing
    # is rewritten, and this is the only thing that keeps a foreign archive
    # (`7-Zip`, `.NET`, or a zip written on Windows) from materialising a
    # file name containing backslashes. Deleting it would be a silent
    # behaviour change on Linux/macOS, so it stays.
    if "\\" in name:
        raise UnsafeArchiveError(f"archive entry is not a POSIX path: {name!r}")
    if name.startswith("/"):
        raise UnsafeArchiveError(f"archive entry is absolute: {name!r}")
    norm = posixpath.normpath(name)
    if norm.startswith("..") or norm == "..":
        raise UnsafeArchiveError(f"archive entry escapes the root: {name!r}")
    last = norm.rsplit("/", 1)[-1]
    # Win32 resolves "NUL", "nul", "CON.txt" and "word/CON" to devices. "NUL"
    # swallows the write and reports success, so a part would silently vanish.
    if last.split(".", 1)[0].strip().upper() in _RESERVED_DEVICE_NAMES:
        raise UnsafeArchiveError(f"archive entry is a device name: {name!r}")
    # Win32 strips a trailing dot or space, so "evil.txt " lands as "evil.txt":
    # two entries collide on one file and the tree stops matching the archive.
    if last != last.rstrip(" ."):
        raise UnsafeArchiveError(
            f"archive entry has a trailing dot or space: {name!r}"
        )
    return norm


def unpack(docx_path: str | Path, out_dir: str | Path) -> Path:
    """Extract every part of ``docx_path`` into ``out_dir``.

    Returns the output directory. Existing contents are left in place (the
    caller usually owns a fresh temp directory).
    """
    src = Path(docx_path)
    dst = Path(out_dir)
    if not src.is_file():
        raise FileNotFoundError(f"not a file: {src}")
    dst.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src) as zf:
        for info in zf.infolist():
            name = _safe_member(info.filename)
            target = dst / name
            # `name` is normpath'd, so a trailing "/" cannot survive it; the old
            # `or name.endswith("/")` half of this test was unreachable. A
            # directory entry is only detectable through `info.is_dir()`, which
            # reads the *unnormalized* name.
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as fsrc, open(target, "wb") as fdst:
                shutil.copyfileobj(fsrc, fdst)
    return dst


def pack(src_dir: str | Path, out_path: str | Path,
         original: str | Path | None = None) -> Path:
    """Zip everything under ``src_dir`` into ``out_path``.

    When ``original`` is given, its entry order and compression method are
    reused, and any part present in the original but missing from ``src_dir`` is
    carried over untouched. This keeps round-trips lossless even if a caller only
    edited one part.

    ``original`` is validated, not sniffed: a path that does not resolve to a
    readable zip archive raises (``ValueError`` for an empty path,
    ``FileNotFoundError`` for a missing/not-a-file path, ``BadZipFile`` for a
    file that is not a zip). Silently packing without it produced a .docx that
    looked fine but had lost every part the caller did not touch. Pass ``None``
    (the default) to opt out of carry-over entirely.
    """
    src = Path(src_dir)
    out = Path(out_path)
    if not src.is_dir():
        raise NotADirectoryError(f"not a directory: {src}")

    present = {
        p.relative_to(src).as_posix()
        for p in src.rglob("*")
        if p.is_file()
    }

    order: list[str] = []
    stored: set[str] = set()

    source = None
    if original is not None:
        # An empty string is the usual CLI-argument default and must not be read
        # as "no original": that is exactly the silent downgrade being closed.
        if str(original).strip() == "":
            raise ValueError(
                "original is an empty path; pass None to pack without carry-over"
            )
        orig = Path(original)
        if not orig.is_file():
            raise FileNotFoundError(f"original archive not found: {orig}")
        try:
            source = zipfile.ZipFile(orig)
        except zipfile.BadZipFile as exc:
            raise zipfile.BadZipFile(
                f"original is not a valid zip archive: {orig}"
            ) from exc
        try:
            for info in source.infolist():
                if info.is_dir():
                    continue
                name = _safe_member(info.filename)
                order.append(name)
                if info.compress_type == zipfile.ZIP_STORED:
                    stored.add(name)
        except BaseException:
            # An unsafe entry aborts the pack before the writing block below, so
            # this is the only place that can close the handle. Without it the
            # original stays open (and locked) on Windows.
            source.close()
            raise

    # Anything new the caller added goes at the end, in a stable order.
    for name in sorted(present - set(order)):
        order.append(name)

    out.parent.mkdir(parents=True, exist_ok=True)
    try:
        with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
            for name in order:
                member = src / name
                method = (zipfile.ZIP_STORED if name in stored
                          else zipfile.ZIP_DEFLATED)
                if member.is_file():
                    zf.write(member, name, compress_type=method)
                elif source is not None and name in source.namelist():
                    # Present in the original but absent from src_dir (most
                    # often a part the caller never touched): carry it over
                    # byte-for-byte so an edit to one part loses nothing.
                    zf.writestr(zipfile.ZipInfo(name), source.read(name),
                                compress_type=method)
    finally:
        if source is not None:
            source.close()
    return out


_USAGE = """\
docx_pack is a library module, not a command-line tool.

Import it instead of running it:

    from docx_pack import unpack, pack

    unpack(docx_path, out_dir) -> Path
        Extract every part of a .docx into out_dir. Rejects entries that would
        escape out_dir (../, absolute paths, Windows drive and UNC prefixes) and
        entries Windows would silently rewrite (device names such as NUL/CON,
        trailing dots or spaces).

    pack(src_dir, out_path, original=None) -> Path
        Zip src_dir into out_path. Pass original= to reuse its entry order and
        compression method, so editing one part keeps the file byte-faithful.
        A missing, empty or non-zip original raises instead of quietly packing
        a document without carry-over.

    from docx_pack import UnsafeArchiveError
        Raised by both functions for entries that are not safe members.

There are no subcommands: any argument other than -h/--help exits 2.
"""


def _main(argv: "list[str]") -> int:
    """Print usage for -h/--help, refuse everything else (no silent success)."""
    if not argv or argv[0] in ("-h", "--help"):
        print(_USAGE, end="")
        return 0
    print(
        "docx_pack is a library module and has no CLI subcommands; "
        "use `from docx_pack import unpack, pack`.",
        file=sys.stderr,
    )
    print("Run `python docx_pack.py --help` for the API summary.",
          file=sys.stderr)
    return 2


if __name__ == "__main__":
    raise SystemExit(_main(sys.argv[1:]))
