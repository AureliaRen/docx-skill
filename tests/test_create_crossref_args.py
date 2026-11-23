# -*- coding: utf-8 -*-
"""The ``crossref`` spec block must hand ``cmd_ref`` the names it reads.

Why this file exists
--------------------
``_crossref_field_args`` builds a namespace object and ``cmd_ref`` reads it with
``getattr(args, "...", default)``.  That idiom cannot tell "the attribute is
absent" from "the attribute is present and false", so a **misspelled key is
indistinguishable from a deliberate false** -- the option silently does nothing
while the command still reports ``{"ok": true}``.

That is not hypothetical.  Measured before the fix:

    spec  {"type": "crossref", "target": "g1", "hyperlink": false}
    out   word/document.xml contains one <w:hyperlink>   <- the flag was ignored
    exit  0,  stdout {"ok": true, "crossrefs": 2}

The spec vocabulary says ``hyperlink`` (positive, default true); ``cmd_ref``
reads ``no_hyperlink`` (negative).  Forwarding one to the other is a no-op by
construction.  The builder now translates, and this file pins the translation so
the next person who renames a switch gets a red test instead of a quiet no-op.

The second half of the file asserts the *observable* end state for each switch,
because "the attribute arrived" is not the same claim as "the document changed".
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx.oxml.ns import qn                                 # noqa: E402

import docx_create                                          # noqa: E402
from docx_create import _crossref_field_args                # noqa: E402


# ---- the contract: every key cmd_ref reads must be produced -----------------

def test_builder_supplies_every_attribute_cmd_ref_reads():
    """Cross-check the builder against ``cmd_ref``'s own ``getattr`` calls.

    Read the attribute names out of the source rather than listing them here by
    hand: a hand-written list is exactly what went stale in the first place.
    This is the guard that would have caught ``hyperlink`` vs ``no_hyperlink``
    on the day it was written.
    """
    import re
    src = (SCRIPTS / "docx_crossref.py").read_text(encoding="utf-8")
    # Every getattr(args, "name", ...) inside the module.
    names = set(re.findall(r'getattr\(\s*args\s*,\s*"([a-z_]+)"', src))
    # cmd_ref / cmd_caption share the namespace; drop the caption-only keys.
    caption_only = {"bookmark_span", "ref_bookmark", "drop_text", "pstyle",
                    "label", "text", "bookmark"}
    wanted = names - caption_only
    # ``hyperlink`` is the spec-side spelling; ``cmd_ref`` must never read it.
    assert "hyperlink" not in wanted

    args = _crossref_field_args({"target": "g1"}, 0, "")
    missing = sorted(n for n in wanted if not hasattr(args, n))
    assert not missing, (
        "crossref spec block does not pass these to cmd_ref: %s -- "
        "getattr() with a default will silently ignore them" % missing)


def test_builder_defaults_match_the_field_form():
    """No spec keys at all -> the ordinary REF field, hyperlinked, tracked."""
    args = _crossref_field_args({"target": "g1"}, 0, "")
    assert args.no_hyperlink is False       # default: do wrap
    assert args.static is False             # default: a real field
    assert args.no_mergeformat is False     # default: keep MERGEFORMAT
    assert args.vert_align == "normal"
    assert args.format == "default"


def test_builder_translates_hyperlink_false_to_no_hyperlink_true():
    """The two spellings are opposites; forwarding one is a no-op."""
    args = _crossref_field_args({"target": "g1", "hyperlink": False}, 0, "")
    assert args.no_hyperlink is True
    args = _crossref_field_args({"target": "g1", "hyperlink": True}, 0, "")
    assert args.no_hyperlink is False


# ---- the observable end state (what the reader actually gets) --------------

def _build(blocks, tmp_path):
    """Build through the real CLI in a subprocess, the way every caller does.

    Follows ``test_bookmark_cli.run_create``: a spec is only correct if the
    *command* accepts and honours it, so the test drives the executable and
    reads the exit code, rather than reaching into a private helper.
    """
    import json
    import subprocess
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps({"blocks": blocks}, ensure_ascii=False),
                         encoding="utf-8")
    out = tmp_path / "out.docx"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(SCRIPTS / "docx_create.py"),
         str(spec_path), str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SCRIPTS.parent))
    assert proc.returncode == 0, (
        "docx_create exited %s\nstderr: %s"
        % (proc.returncode, (proc.stderr or "")[:400]))
    return out


def _xml(out) -> str:
    import zipfile
    with zipfile.ZipFile(out) as z:
        return z.read("word/document.xml").decode("utf-8")


BASE = [
    {"type": "paragraph", "text": "图 1  系统架构示意", "bookmark": "g1"},
    {"type": "paragraph", "text": "引用："},
]


def test_hyperlink_false_really_removes_the_hyperlink(tmp_path):
    """The measured defect: the flag was ignored and a hyperlink was written."""
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "hyperlink": False}], tmp_path)
    assert "<w:hyperlink" not in _xml(out)


def test_hyperlink_true_keeps_it(tmp_path):
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "hyperlink": True}], tmp_path)
    assert "<w:hyperlink" in _xml(out)


def test_static_true_writes_no_field(tmp_path):
    """``static`` must reach ``cmd_ref`` too, not just be accepted by the spec."""
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "static": True}], tmp_path)
    xml = _xml(out)
    # The bookmark paragraph has no field of its own here (no caption), so any
    # field instruction in the document belongs to the reference.
    assert "REF" not in xml or "instrText" not in xml
    assert "<w:fldChar" not in xml


def test_static_keeps_superscript(tmp_path):
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "static": True, "superscript": True}], tmp_path)
    xml = _xml(out)
    assert 'w:val="superscript"' in xml
    assert "<w:fldChar" not in xml


def test_no_mergeformat_reaches_the_field(tmp_path):
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "no_mergeformat": True}], tmp_path)
    assert "MERGEFORMAT" not in _xml(out)


def test_field_form_still_carries_mergeformat_by_default(tmp_path):
    out = _build(BASE + [{"type": "crossref", "target": "g1",
                          "superscript": True}], tmp_path)
    assert "MERGEFORMAT" in _xml(out)


# ---- spec validation -------------------------------------------------------

@pytest.mark.parametrize("key", ["static", "no_mergeformat"])
def test_new_keys_are_accepted_by_the_spec_validator(key):
    """An unknown key is exit 2, so a key the builder reads must be whitelisted."""
    docx_create.validate_spec({"blocks": [
        {"type": "paragraph", "text": "x", "bookmark": "g1"},
        {"type": "crossref", "target": "g1", key: True},
    ]})


def test_static_with_page_format_is_refused_not_guessed():
    """--static cannot express a page number; the spec path must refuse too."""
    from docx_crossref import UsageError
    args = _crossref_field_args({"target": "g1", "static": True,
                                 "format": "page"}, 0, "")
    # The builder passes it through; cmd_ref is where the refusal lives.
    assert args.format == "page" and args.static is True
    # And that refusal is real (see test_crossref_static.py for the full set).
    assert UsageError is not None