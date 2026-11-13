# -*- coding: utf-8 -*-
"""``image`` can be sized as a share of the text width, not only in millimetres.

The reported problem: an absolute ``width_mm`` is the one number an author cannot
know without measuring the page.  "Full text width" is what they mean, and it is
158 mm on A4 with 26 mm side margins — a number that changes with the margins, so
hard-coding it just relocates the defect.  ``width_frac`` resolves against the
*current section* via ``content_width_mm``, the helper the TOC tab stop already
uses for exactly that reason.

Two failure modes are pinned here because both are the repository's recurring
kind — a silent choice between two given answers, and a value quietly clamped
into range:

* ``width_frac`` together with ``width_mm`` is a contradiction, so it is exit 2
  rather than "one of them wins";
* ``0.0`` and ``1.5`` are refused with the legal range, never clamped.

The truthiness trap is pinned separately: ``if block.get("width_frac")`` reads
``0.0`` as "not given" and would fall back to the picture's intrinsic size.
"""
from __future__ import annotations

import json
import struct
import subprocess
import sys
import zlib
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL / "scripts"
CREATE = SCRIPTS / "docx_create.py"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402

import docx_create as dc  # noqa: E402

A4_26 = {"width_mm": 210, "height_mm": 297,
         "margins_mm": {"top": 26, "bottom": 26, "left": 26, "right": 26}}
A4_20 = {"width_mm": 210, "height_mm": 297,
         "margins_mm": {"top": 25, "bottom": 25, "left": 20, "right": 20}}

# A4 with 26 mm side margins is 210 - 26 - 26 = 158 mm of text width; with 20 mm
# margins it is 170 mm.  Both are asserted so the fraction is shown to follow the
# geometry rather than to equal the one number from the report.
WIDTH_158 = 158.0
WIDTH_170 = 170.0
TOL = 0.5


# ---- helpers -----------------------------------------------------------------

def make_png(path: Path, w=2, h=2) -> None:
    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))
    ihdr = struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    path.write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
                     + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


@pytest.fixture()
def png(tmp_path):
    path = tmp_path / "pic.png"
    make_png(path)
    return path


def run_create(tmp_path, spec):
    spec_path = tmp_path / "spec.json"
    spec_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
    out = tmp_path / "out.docx"
    proc = subprocess.run(
        [sys.executable, "-X", "utf8", str(CREATE), str(spec_path), str(out)],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        cwd=str(SKILL))
    payload = None
    for line in (proc.stdout or "").splitlines():
        if line.strip().startswith("{"):
            payload = json.loads(line)
    return proc.returncode, payload, proc.stderr or "", out


def picture_mm(path: Path) -> float:
    doc = Document(str(path))
    assert doc.inline_shapes, "no picture in the document"
    return doc.inline_shapes[0].width.mm


# ---- 1. the two numbers the report asks for ----------------------------------

def test_width_frac_one_is_the_full_text_width(tmp_path, png):
    """1.0 on A4 with 26 mm side margins must measure 158 mm (±0.5)."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 1.0}]})
    assert rc == 0, stderr
    assert picture_mm(out) == pytest.approx(WIDTH_158, abs=TOL)


def test_width_frac_half_is_half_of_it(tmp_path, png):
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 0.5}]})
    assert rc == 0, stderr
    assert picture_mm(out) == pytest.approx(WIDTH_158 / 2, abs=TOL)


def test_the_fraction_follows_the_margins_not_a_constant(tmp_path, png):
    """Same 1.0, different margins ⇒ 170 mm, not 158 mm.

    This is the assertion that separates "resolve against the section" from
    "hard-code 158"; the latter passes the first test and fails this one.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_20, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 1.0}]})
    assert rc == 0, stderr
    assert picture_mm(out) == pytest.approx(WIDTH_170, abs=TOL)


@pytest.mark.parametrize("frac,share", [(1.0, 1.0), (0.75, 0.75), (0.5, 0.5),
                                        (0.25, 0.25), (0.1, 0.1)])
def test_fractions_scale_linearly(tmp_path, png, frac, share):
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": frac}]})
    assert rc == 0, stderr
    assert picture_mm(out) == pytest.approx(WIDTH_158 * share, abs=TOL)


def test_the_resolved_width_matches_content_width_mm(tmp_path, png):
    """Cross-check against the helper rather than against a literal.

    ``content_width_mm`` already rounds to 0.01 mm and already exists for the
    TOC tab stop; reusing it is the whole point, so the fraction must agree with
    it to within the same tolerance.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 0.5}]})
    assert rc == 0, stderr
    doc = Document(str(out))
    expected = dc.content_width_mm(doc.sections[-1]) * 0.5
    assert picture_mm(out) == pytest.approx(expected, abs=TOL)


# ---- 2. width_mm still works, and the two are exclusive ---------------------

def test_width_mm_still_wins_when_it_is_the_only_key(tmp_path, png):
    """Backward compatibility: the absolute key is untouched."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_mm": 60}]})
    assert rc == 0, stderr
    assert picture_mm(out) == pytest.approx(60.0, abs=TOL)


def test_neither_key_leaves_the_intrinsic_size(tmp_path, png):
    """A 2x2 PNG at 96 dpi is tiny; the point is only that no width was forced."""
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png)}]})
    assert rc == 0, stderr
    assert picture_mm(out) < 10, "a 2px image should not be scaled up"


def test_width_frac_and_width_mm_together_are_refused(tmp_path, png):
    """Exit 2, one ``error:`` line, and no file — the D-024 shape."""
    rc, payload, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 0.5,
         "width_mm": 60}]})
    assert rc == 2, (rc, stderr)
    assert payload is not None and payload["ok"] is False
    assert payload["output"] is None
    assert "mutually exclusive" in payload["error"]
    assert "width_frac" in payload["error"] and "width_mm" in payload["error"]
    assert not out.exists(), "exit 2 promises nothing was written"
    lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    assert lines[0].startswith("error: ")
    assert "Traceback" not in stderr


# ---- 3. out of range: refused with the range, never clamped ----------------

@pytest.mark.parametrize("bad", [0.0, -0.5, 1.5, 2, 100])
def test_out_of_range_fractions_are_refused(tmp_path, png, bad):
    rc, payload, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": bad}]})
    assert rc == 2, (bad, rc, stderr)
    assert payload["ok"] is False
    assert "out of range" in payload["error"]
    # The range is spelled out, and clamping is explicitly disclaimed.
    assert "0 < f <= 1" in payload["error"]
    assert "not clamped" in payload["error"]
    assert not out.exists()


def test_width_frac_zero_is_not_treated_as_missing(tmp_path, png):
    """The truthiness trap, stated as its own test.

    ``if block.get("width_frac"):`` would read 0.0 as "no width given" and build
    a document at the picture's intrinsic size with ``exit 0`` — a silent success
    for a spec that is wrong.  The expected answer is a refusal.
    """
    rc, payload, _, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": 0.0}]})
    assert rc == 2, "0.0 must be a range error, not a silent fallback"
    assert "out of range" in payload["error"]
    assert not out.exists()


@pytest.mark.parametrize("bad", ["half", True, [0.5], {"f": 0.5}])
def test_a_non_number_fraction_is_refused(tmp_path, png, bad):
    """``True`` is excluded explicitly: ``bool`` is an ``int`` in Python."""
    rc, payload, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": bad}]})
    assert rc == 2, (bad, rc, stderr)
    assert "width_frac" in payload["error"]
    assert not out.exists()


def test_json_null_is_read_as_not_given(tmp_path, png):
    """``"width_frac": null`` is absence, not zero — and it does not crash.

    ``block.get()`` cannot tell an explicit ``null`` from an omitted key, so
    ``null`` takes the same path as leaving the key out: the image keeps its
    intrinsic size and the run succeeds.  Pinned so the behaviour is a decision
    rather than an accident, and so a future reader does not "fix" it into a
    crash by adding an unconditional ``float(block["width_frac"])``.
    """
    rc, _, stderr, out = run_create(tmp_path, {"page": A4_26, "blocks": [
        {"type": "image", "path": str(png), "width_frac": None}]})
    assert rc == 0, stderr
    assert picture_mm(out) < 10, "null must not become a 158 mm figure"


# ---- 4. the validator, without going through the CLI -----------------------

def test_the_checker_is_the_validator_not_the_builder():
    """Range/exclusivity live in ``validate_spec``: rejected before any write."""
    problems: list = []
    dc._check_width_frac("blocks[0]", {"width_frac": 1.5}, problems)
    assert problems and "out of range" in problems[0]

    problems = []
    dc._check_width_frac("blocks[0]", {"width_frac": 0.5, "width_mm": 10},
                         problems)
    assert problems and "mutually exclusive" in problems[0]

    problems = []
    dc._check_width_frac("blocks[0]", {"width_frac": 0.5}, problems)
    assert problems == []

    problems = []
    dc._check_width_frac("blocks[0]", {"width_mm": 10}, problems)
    assert problems == [], "width_mm alone must not be touched"


def test_both_keys_is_reported_once_not_twice():
    """Exclusivity outranks the range check: one sentence, not two."""
    problems: list = []
    dc._check_width_frac("blocks[0]", {"width_frac": 9, "width_mm": 1},
                         problems)
    assert len(problems) == 1, problems
    assert "mutually exclusive" in problems[0]
