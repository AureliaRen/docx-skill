# -*- coding: utf-8 -*-
"""A rejected spec must say what shape it expected, not only what it saw.

The reported problem (P3): ``"margins_mm": [25, 25, 20, 20]`` produced
``must be a JSON object, got list``, and ``{"margin_top": 25}`` produced
``unknown key "margin_top"; valid keys here: …``.  Both are *correct* and neither
is *usable*: the first does not name the four keys, and the second lists the keys
flat, which is exactly the part a nested value got wrong.

This file is **message text only**.  No assertion here changes what is accepted
or rejected — the vocabulary and the rules are pinned elsewhere
(``test_docx_skill.py``), and the ``unknown key … valid keys here`` sentence is
kept verbatim with the shape appended after it, so its existing tests still pass.
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
sys.path.insert(0, str(SCRIPTS))

import docx_create as dc  # noqa: E402

MARGINS_EXAMPLE = '{"top": 25, "bottom": 25, "left": 20, "right": 20}'


# ---- helpers -----------------------------------------------------------------

def run_create(tmp_path, spec):
    spec_path = tmp_path / "spec.json"
    if isinstance(spec, str):
        spec_path.write_text(spec, encoding="utf-8")
    else:
        spec_path.write_text(json.dumps(spec, ensure_ascii=False),
                             encoding="utf-8")
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


def error_of(tmp_path, spec) -> str:
    rc, payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 2, (rc, stderr)
    assert payload is not None and payload["ok"] is False
    assert not out.exists(), "a refused spec must not leave a file"
    lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    assert lines[0].startswith("error: "), lines[0]
    assert "Traceback" not in stderr
    return payload["error"]


# ---- 1. the list-shaped margins_mm -----------------------------------------

def test_margins_as_a_list_names_the_four_keys(tmp_path):
    """The exact mistake from the report: right numbers, wrong container."""
    error = error_of(tmp_path, {"page": {"width_mm": 210, "height_mm": 297,
                                         "margins_mm": [25, 25, 20, 20]},
                                "blocks": []})
    assert "margins_mm" in error
    assert "must be a JSON object, got list" in error
    assert MARGINS_EXAMPLE in error, error


def test_the_margins_example_also_names_the_optional_keys(tmp_path):
    """``header`` / ``footer`` / ``gutter`` are legal, so the hint says so."""
    error = error_of(tmp_path, {"page": {"margins_mm": [1, 2, 3, 4]},
                                "blocks": []})
    for optional in ("header", "footer", "gutter"):
        assert optional in error, error


def test_a_margin_value_of_the_wrong_type_is_out_of_scope(tmp_path):
    """``{"top": [25]}`` has a *valid* key, so it is not a key problem at all.

    It reaches the builder and fails there.  Pinning the current answer keeps the
    boundary honest: P3 was about the *message* for a wrong container, and the
    key vocabulary is untouched, so a wrong *value* still takes the builder's
    path (exit 2 via ``_BUILD_ERRORS``) rather than growing a validator branch
    that was never asked for.
    """
    rc, payload, _, out = run_create(tmp_path, {
        "page": {"margins_mm": {"top": [25]}}, "blocks": []})
    assert rc == 2
    assert payload["ok"] is False
    assert not out.exists()


# ---- 2. the misspelt nested key --------------------------------------------

def test_a_typo_inside_margins_mm_shows_the_shape_as_well(tmp_path):
    """``valid keys here`` alone lists names flat; the shape shows the nesting."""
    error = error_of(tmp_path, {"page": {"width_mm": 210, "height_mm": 297,
                                         "margins_mm": {"margin_top": 25}},
                                "blocks": []})
    assert '"margin_top"' in error
    assert "valid keys here" in error            # the existing sentence survives
    assert "expected shape" in error
    assert MARGINS_EXAMPLE in error, error


def test_the_existing_wording_is_unchanged_up_to_the_appended_shape(tmp_path):
    """The prefix is byte-identical to what the repository already promises.

    ``"tp"`` is close enough to ``"top"`` to earn a suggestion, so the expected
    prefix includes it — the point of the assertion is that nothing *before*
    ``; expected shape:`` was rewritten to make room for the new clause.
    """
    error = error_of(tmp_path, {"page": {"margins_mm": {"tp": 25}}, "blocks": []})
    prefix = ('spec.page.margins_mm: unknown key "tp" (did you mean "top"?); '
              "valid keys here: bottom, footer, gutter, header, left, right, "
              "top; expected shape: ")
    assert error.startswith(prefix), error


def test_the_top_level_key_list_is_not_given_a_margins_shape(tmp_path):
    """The shape is scoped: a typo elsewhere must not claim to be about margins."""
    error = error_of(tmp_path, {"blockz": [], "blocks": []})
    assert 'did you mean "blocks"' in error
    assert "valid keys here" in error
    assert MARGINS_EXAMPLE not in error, error


def test_other_nested_objects_still_get_their_own_example(tmp_path):
    """Generalisation check: the hint is per-place, not globally one string."""
    rc, payload, stderr, _ = run_create(tmp_path, {"blocks": [
        {"type": "paragraph", "runs": ["not an object"]}]})
    assert rc == 2, stderr
    assert "expected e.g." in payload["error"]
    assert MARGINS_EXAMPLE not in payload["error"], (
        "a bad run must not be told about page margins")


# ---- 3. non-object blocks and the top-level spec ---------------------------

def test_a_block_that_is_a_list_says_what_a_block_looks_like(tmp_path):
    error = error_of(tmp_path, {"blocks": [["type", "paragraph"]]})
    assert "blocks[0]" in error
    assert "must be a JSON object, got list" in error
    assert "expected e.g." in error, error


def test_a_root_that_is_a_list_is_still_rejected(tmp_path):
    """``spec`` itself is checked before anything else and keeps its message."""
    error = error_of(tmp_path, "[]")
    assert "must be a JSON object" in error
    assert "got list" in error


# ---- 4. the hint function itself -----------------------------------------

@pytest.mark.parametrize("where,must_contain", [
    ("spec", "blocks"),
    ("spec.page.margins_mm", MARGINS_EXAMPLE),
    ("blocks[3].page.margins_mm", MARGINS_EXAMPLE),
    ("spec.page", "blocks"),
    ("blocks[0].runs[2]", "text"),
    ("blocks[0].tab_stops[1]", "pos_mm"),
    ("blocks[0].bottom_border", "single"),
    ("styles[1]", "name"),
    ("blocks[0].items[4]", "text"),
    ("blocks[9]", "type"),
])
def test_shape_hint_covers_every_place_that_expects_an_object(where,
                                                              must_contain):
    """Every nesting level a spec can get wrong has an example of its own."""
    assert must_contain in dc._shape_hint(where)


def test_shape_hint_starts_with_an_object_literal():
    """Whatever the place, the answer opens with an example object.

    It is not required to *end* with one: the ``margins_mm`` example carries a
    trailing note about the three optional keys, which is the only way to say
    "these four are required, those three are not" in one line.
    """
    for where in ("spec", "blocks[0]", "spec.page", "spec.page.margins_mm",
                  "styles[0]", "blocks[0].runs[0]", "weird"):
        hint = dc._shape_hint(where)
        assert hint.startswith("{"), hint
        assert hint.count("{") >= hint.count("}") - 1, hint


# ---- 5. nothing here changed a decision -----------------------------------

def test_a_valid_margins_object_still_builds(tmp_path):
    rc, payload, stderr, out = run_create(tmp_path, {
        "page": {"width_mm": 210, "height_mm": 297,
                 "margins_mm": {"top": 25, "bottom": 25, "left": 20,
                                "right": 20}},
        "blocks": [{"type": "paragraph", "text": "正文"}]})
    assert rc == 0, stderr
    assert out.exists()
    assert payload["ok"] is True


def test_the_optional_margin_keys_are_still_accepted(tmp_path):
    """``header``/``footer``/``gutter`` are in the vocabulary, and stay there."""
    rc, _, stderr, out = run_create(tmp_path, {
        "page": {"margins_mm": {"top": 25, "bottom": 25, "left": 20,
                                "right": 20, "header": 15, "footer": 15,
                                "gutter": 0}},
        "blocks": []})
    assert rc == 0, stderr


def test_the_unknown_key_error_still_lists_sorted_valid_keys(tmp_path):
    """Text-only change: the key list keeps its order and completeness."""
    error = error_of(tmp_path, {"page": {"margins_mm": {"bogus": 1}},
                                "blocks": []})
    listed = error.split("valid keys here: ")[1].split(";")[0]
    keys = [k.strip() for k in listed.split(",")]
    assert keys == sorted(dc.MARGIN_KEYS), keys
