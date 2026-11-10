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
    if "\\" in name:
        raise UnsafeArchiveError(f"archive entry is not a POSIX path: {name!r}")
    if name.startswith("/"):
        raise UnsafeArchiveError(f"archive entry is absolute: {name!r}")
    norm = posixpath.normpath(name)
    if norm.startswith("..") or norm == "..":
        raise UnsafeArchiveError(f"archive entry escapes the root: {name!r}")
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
            if info.is_dir() or name.endswith("/"):
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
        orig = Path(original)
        if orig.is_file():
            source = zipfile.ZipFile(orig)
            for info in source.infolist():
                if info.is_dir():
                    continue
                name = _safe_member(info.filename)
                order.append(name)
                if info.compress_type == zipfile.ZIP_STORED:
                    stored.add(name)

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
        escape out_dir (../, absolute paths, Windows drive letters).

    pack(src_dir, out_path, original=None) -> Path
        Zip src_dir into out_path. Pass original= to reuse its entry order and
        compression method, so editing one part keeps the file byte-faithful.

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
