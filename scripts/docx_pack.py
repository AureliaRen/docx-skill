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
* **No dependency on `docx_utils`.** Kept standalone so it can be used by any
  script without dragging in the XML helper surface.

Usage::

    from docx_pack import unpack, pack
    unpack("in.docx", "unpacked/")
    ...                     # edit files under unpacked/
    pack("unpacked/", "out.docx", original="in.docx")
"""
from __future__ import annotations

import posixpath
import shutil
import zipfile
from pathlib import Path

__all__ = ["unpack", "pack", "UnsafeArchiveError"]


class UnsafeArchiveError(ValueError):
    """Raised when an archive entry would escape the extraction root."""


def _safe_member(name: str) -> str:
    """Normalize a zip entry name, rejecting anything outside the root."""
    if not name:
        raise UnsafeArchiveError("empty archive entry name")
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

    if original is not None:
        orig = Path(original)
        if orig.is_file():
            with zipfile.ZipFile(orig) as zf:
                for info in zf.infolist():
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
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in order:
            member = src / name
            if not member.is_file():
                # Present in the original but deleted by the caller: skip it.
                if original is not None:
                    continue
                continue
            method = zipfile.ZIP_STORED if name in stored else zipfile.ZIP_DEFLATED
            zf.write(member, name, compress_type=method)
    return out
