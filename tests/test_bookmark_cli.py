# -*- coding: utf-8 -*-
"""The bookmark-name guard as the *CLI* sees it: a refusal, not a document.

Every other test in this file family works on functions.  That leaves the one
thing an operator actually depends on untested: ``docx_create.py`` must exit
non-zero and **write no file** when a spec would produce a duplicate or
unnameable bookmark.  A mutation that deletes the whole branch leaves the
function-level tests green and only this file goes red
(work/REVIEW-ADVERSARIAL-R6.md 第三轮 M4).

The three assertions that matter, in order of how hard they are to fake:

1. ``returncode == 1`` — the branch exists at all;
2. the error payload names the *new* failure mode, not the older ``missing``
   one, so a passing test cannot be explained by a different code path;
3. ``not output.exists()`` — the document was never written.  This is the one a
   branch-deleting mutation cannot survive: without the guard, the run succeeds
   and the file lands on disk.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL / "scripts"
CREATE = SCRIPTS / "docx_create.py"


def run_create(tmp_path, spec: dict):
    """Run ``docx_create.py`` in a subprocess; return (rc, payload, output)."""
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.docx"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(CREATE), str(spec_path), str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SKILL))
    payload = None
    for line in (proc.stdout or "").splitlines():
        line = line.strip()
        if line.startswith("{"):
            try:
                payload = json.loads(line)
            except json.JSONDecodeError:
                continue
    return proc.returncode, payload, out


def test_duplicate_bookmark_names_abort_the_build(tmp_path):
    """Two paragraphs claiming the same bookmark name must not produce a file.

    Word silently merges the pair (the later Start wins, so a TOC entry jumps to
    the wrong chapter) and LibreOffice renames the second one, killing every
    field written against the original name.  Neither is visible in the XML we
    wrote, so the only defence is to refuse.
    """
    spec = {"blocks": [
        {"type": "paragraph", "text": "第一章", "bookmark": "same"},
        {"type": "paragraph", "text": "第二章", "bookmark": "same"},
    ]}
    rc, payload, out = run_create(tmp_path, spec)

    assert rc == 1, payload
    assert payload is not None, "the CLI must report why it refused"
    assert payload["error"] == "invalid or duplicate bookmark names"
    assert payload["duplicate_names"] == ["same"]
    assert not out.exists(), "a refused build must not leave a document behind"


def test_an_over_long_bookmark_name_aborts_the_build(tmp_path):
    spec = {"blocks": [
        {"type": "paragraph", "text": "第一章", "bookmark": "A" * 41},
    ]}
    rc, payload, out = run_create(tmp_path, spec)

    assert rc == 1, payload
    assert payload["error"] == "invalid or duplicate bookmark names"
    assert payload["invalid_names"][0]["name"] == "A" * 41
    assert any("40" in problem for problem in payload["invalid_names"][0]["problems"])
    assert not out.exists()


def test_a_space_in_a_bookmark_name_aborts_the_build(tmp_path):
    """Word and LibreOffice both rewrite the space to ``_`` on save."""
    spec = {"blocks": [
        {"type": "paragraph", "text": "第一章", "bookmark": "_Ref图 1"},
    ]}
    rc, payload, out = run_create(tmp_path, spec)

    assert rc == 1, payload
    assert payload["invalid_names"][0]["name"] == "_Ref图 1"
    assert not out.exists()


def test_a_punctuated_bookmark_name_is_written(tmp_path):
    """The counterpart: names Word accepts must not be refused.

    Without this, a guard that rejected everything would pass the three tests
    above.
    """
    spec = {"blocks": [
        {"type": "paragraph", "text": "第一章", "bookmark": "图(1)"},
        {"type": "paragraph", "text": "第二章", "bookmark": "图（2）"},
    ]}
    rc, payload, out = run_create(tmp_path, spec)

    assert rc == 0, payload
    assert payload["ok"] is True
    assert out.exists()
