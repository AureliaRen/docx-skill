# -*- coding: utf-8 -*-
"""Forward cross-references: a body citation whose target comes later.

The owner's acceptance shape is a paper: the body says "参见文献[1]" and the
reference list owning ``ref_1`` sits at the end. When the body is written the
bookmark therefore does not exist yet. ``cmd_ref`` used to raise
``SystemExit("no such bookmark")`` on that, which made the single most common
citation shape impossible to express through a spec.

What is pinned here:
  * ``cmd_ref`` writes the field anyway when told to (``allow_missing``);
  * ``docx_create`` always passes it, because a spec is written top-down;
  * a *dangling* target is still an error, reported after the whole document
    exists, rather than a silently dead click.

XML level and ``tmp_path`` based: no Word needed, so CI covers it.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document                                   # noqa: E402
from docx.oxml.ns import qn                                 # noqa: E402

from docx_crossref import (                                 # noqa: E402
    all_bookmarks, check_crossref_anchors, cmd_ref,
)
from docx_numbering import cmd_define, set_numpr              # noqa: E402


class Args:
    def __init__(self, **kw):
        self.__dict__.update(kw)


def _ref_args(**kw):
    base = dict(index=0, target="ref_1", prefix="", suffix="",
                format="paragraph-number", cached=None, vert_align="normal",
                superscript=False, hyperlink=True, allow_missing=True)
    base.update(kw)
    return Args(**base)


def _doc_with_body_and_refs():
    """Body paragraph first, then the numbered reference list (``ref_1``..)."""
    d = Document()
    body = d.add_paragraph("参见文献")
    nid = cmd_define(d, Args(text="[%1]", fmt="decimal", start=1, jc="left",
                             indent_left=425, hanging=425))["numId"]
    from docx_crossref import bm_name, _bookmark_pair
    from docx_utils import next_bm_id
    for k in (1, 2):
        p = d.add_paragraph("文献 %d" % k)
        set_numpr(p, nid, 0)
        bid = next_bm_id(d.element.body)
        start, end = _bookmark_pair("ref_%d" % k, bid)
        p._p.insert(0, start)
        p._p.append(end)
    return d, body


def test_forward_ref_writes_without_target_present():
    """The bookmark does not exist yet, and that must not be fatal."""
    d, body = _doc_with_body_and_refs()
    # drop the bookmarks to model "body written before the reference list"
    for el in list(d.element.body.iter(qn("w:bookmarkStart"))):
        el.getparent().remove(el)
    assert not all_bookmarks(d)

    out = cmd_ref(d, _ref_args())
    assert out["ok"] is True
    assert out["target"] == "ref_1"
    # the field really landed, wrapped in a clickable anchor
    instr = [i.text for i in d.element.body.iter(qn("w:instrText"))]
    assert any("REF ref_1" in (t or "") for t in instr)
    anchors = [h.get(qn("w:anchor")) for h in d.element.body.iter(qn("w:hyperlink"))]
    assert "ref_1" in anchors


def test_unknown_target_still_fatal_without_allow_missing():
    """The guard is a default, not removed: a typo must still be caught."""
    d = Document()
    d.add_paragraph("正文")
    with pytest.raises(SystemExit):
        cmd_ref(d, _ref_args(allow_missing=False))


def test_check_anchors_reports_dangling_target():
    d, _body = _doc_with_body_and_refs()
    cmd_ref(d, _ref_args())
    report = check_crossref_anchors(d)
    assert report["missing"] == []
    assert report["resolved"] >= 1

    # remove every bookmark: the reference is now dangling
    for el in list(d.element.body.iter(qn("w:bookmarkStart"))):
        el.getparent().remove(el)
    report = check_crossref_anchors(d)
    assert "ref_1" in report["missing"]


SPEC = {
    "blocks": [
        {"type": "paragraph", "text": "正文引用参考文献，见文献"},
        {"type": "crossref", "target": "ref_1", "format": "paragraph-number",
         "superscript": True},
        {"type": "page_break"},
        {"type": "numbered_list", "num_text": "[%1]", "bookmark_prefix": "ref_",
         "items": ["A 文献", "B 文献"]},
    ]
}


def test_create_spec_forward_reference_end_to_end(tmp_path: Path):
    """A spec written top-down resolves its own forward references."""
    spec = tmp_path / "spec.json"
    spec.write_text(json.dumps(SPEC, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.docx"
    r = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_create.py"), str(spec), str(out)],
        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 0, r.stderr
    payload = json.loads(r.stdout)
    assert payload["ok"] is True
    assert payload["crossrefs"] >= 1

    doc = Document(str(out))
    assert "ref_1" in all_bookmarks(doc)
    assert check_crossref_anchors(doc)["missing"] == []
    anchors = [h.get(qn("w:anchor")) for h in doc.element.body.iter(qn("w:hyperlink"))]
    assert "ref_1" in anchors


def test_create_spec_dangling_target_fails_loudly(tmp_path: Path):
    """A spec that points at nothing must fail, not ship a dead click."""
    bad = {"blocks": [
        {"type": "paragraph", "text": "正文"},
        {"type": "crossref", "target": "nope_9", "format": "default"},
    ]}
    spec = tmp_path / "bad.json"
    spec.write_text(json.dumps(bad, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "bad.docx"
    r = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_create.py"), str(spec), str(out)],
        capture_output=True, text=True, encoding="utf-8")
    assert r.returncode == 1
    payload = json.loads(r.stdout)
    assert payload["ok"] is False
    assert "nope_9" in payload["missing"]
    assert not out.exists()
