# -*- coding: utf-8 -*-
"""``template="zh"`` and ``custom``'s ``{page}`` / ``{total}`` placeholders.

The Chinese footer ``第 X 页　共 Y 页`` was reachable only by hand-writing
PAGE/NUMPAGES fields through ``postprocess.py`` after the fact; ``page_of``
hard-coded an English ``"Page "`` / ``" of "`` and ``custom`` had a PAGE field
with no NUMPAGES at all, so "共 Y 页" could not be expressed.  These tests pin
the three things that make the feature real rather than nominal:

* the **order** of the five segments (text / field / text / field / text) and
  the fact that the total field is a *field*, not the literal string "1";
* the separator between 页 and 共 is **U+3000**, a full-width space, which is
  typographically the right one and visually indistinguishable from U+0020 in
  most editors -- hence the codepoint assertion, not a string comparison;
* every pre-existing template still lays out exactly as before, which is what
  lets ``custom`` grow placeholders without changing any old document.

Assertions read the saved part XML rather than ``Paragraph.text``: a footer
that is structurally wrong still opens and still prints something.
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402
import docx_section as ds  # noqa: E402

FULLWIDTH_SPACE = "\u3000"


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def make_footer(tmp_path, **kwargs):
    """Run the function on a fresh document; return (result, footer part)."""
    doc = Document()
    doc.add_paragraph("body")
    result = ds.add_page_number_field(doc.sections[-1], **kwargs)
    path = tmp_path / "f.docx"
    doc.save(path)
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/footer1.xml").decode("utf-8")
    return result, xml, doc


def paragraph_xml(xml: str) -> str:
    start = xml.index("<w:p>")
    return xml[start:xml.index("</w:p>") + 6]


def _iter_text(xml: str, tag: str) -> list:
    open_tag = "<%s" % tag
    out = []
    pos = 0
    while True:
        i = xml.find(open_tag, pos)
        if i < 0:
            return out
        gt = xml.index(">", i)
        end = xml.index("</%s>" % tag, gt)
        out.append(xml[gt + 1:end])
        pos = end
    return out


def texts(xml: str) -> list:
    return _iter_text(xml, "w:t")


def instrs(xml: str) -> list:
    return [t.strip() for t in _iter_text(xml, "w:instrText")]


def body(xml: str) -> str:
    """The paragraph element only (drops the <w:ftr> namespace soup)."""
    return paragraph_xml(xml)


def text_runs(xml: str) -> list:
    """The ``<w:t>`` texts that sit *outside* every field."""
    frag = body(xml)
    out = []
    depth = 0
    pos = 0
    while True:
        i = min((x for x in (frag.find("<w:fldChar", pos),
                             frag.find("<w:t", pos)) if x >= 0), default=-1)
        if i < 0:
            return out
        if frag.startswith("<w:fldChar", i):
            kind = frag.index('w:fldCharType="', i) + len('w:fldCharType="')
            fld = frag[kind:frag.index('"', kind)]
            depth += 1 if fld == "begin" else (-1 if fld == "end" else 0)
            pos = frag.index("/>", i) + 2
        else:
            gt = frag.index(">", i)
            end = frag.index("</w:t>", gt)
            if depth == 0:
                out.append(frag[gt + 1:end])
            pos = end


def run_footer_cli(tmp_path, *argv):
    """``docx_section.py pagenums`` on a seeded document; return (proc, out)."""
    seed = tmp_path / "seed.docx"
    d = Document()
    d.add_paragraph("body")
    d.save(seed)
    out = tmp_path / "out.docx"
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_section.py"), "pagenums",
         str(seed), "-o", str(out)] + list(argv),
        capture_output=True, check=False,
        env=dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8"))
    return proc, out


def run_create_cli(tmp_path, spec, name="spec.json"):
    spec_path = tmp_path / name
    spec_path.write_text(json.dumps(spec, ensure_ascii=False),
                         encoding="utf-8")
    out = tmp_path / (name + ".docx")
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_create.py"),
         str(spec_path), str(out)],
        capture_output=True, check=False,
        env=dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8"))
    return proc, out


def footer_xml(path) -> str:
    with zipfile.ZipFile(path) as z:
        return z.read("word/footer1.xml").decode("utf-8")


# ---------------------------------------------------------------------------
# 1. the zh layout: segment order, the two fields, the full-width space
# ---------------------------------------------------------------------------
def test_zh_template_lays_out_five_segments_in_order(tmp_path):
    """第 / PAGE / 页　共 / NUMPAGES / 页 -- nothing dropped, nothing reordered."""
    _res, xml, _doc = make_footer(tmp_path, template="zh")
    frag = body(xml)

    # The literal order of the run children is the whole contract: a footer
    # reading "第 页 共 1 页 1" would still contain every piece.
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    outside = text_runs(xml)
    assert outside == ["第 ", " 页" + FULLWIDTH_SPACE + "共 ", " 页"], outside
    # ...and in that physical order, interleaved with the fields.
    assert _child_runs(frag) == ["t", "f", "t", "f", "t"]


def _child_runs(frag: str) -> list:
    """Classify every ``<w:r>`` of the paragraph as 't' (text) or 'f' (field).

    A field's five runs (begin / instrText / separate / result / end) are one
    logical segment, so they collapse to a single ``'f'``.  The *result* run is
    a plain ``<w:t>``, so a naive text/field test would split the field in two:
    the ``begin``/``separate``/``end`` depth is what decides.
    """
    out = []
    depth = 0
    pos = 0
    while True:
        i = frag.find("<w:r>", pos)
        if i < 0:
            return out
        j = frag.find("</w:r>", i)
        chunk = frag[i:j]
        pos = j + 6
        opened = chunk.count('w:fldCharType="begin"')
        closed = chunk.count('w:fldCharType="end"')
        inside = depth > 0 or opened
        if inside and not (out and out[-1] == "f"):
            out.append("f")
        elif not inside and "<w:t" in chunk:
            out.append("t")
        depth += opened - closed


def test_zh_separator_is_u3000_not_ascii_space(tmp_path):
    """The gap between 页 and 共 is a full-width space.

    A half-width space here is invisible in a diff and wrong on the page, so
    the assertion is on the codepoint, and it also states what it must *not*
    be.
    """
    _res, xml, _doc = make_footer(tmp_path, template="zh")
    middle = text_runs(xml)[1]
    assert middle == " 页" + FULLWIDTH_SPACE + "共 "
    assert "\u3000" in middle
    assert middle.count("\u3000") == 1
    # No ASCII space in the two places the full-width one belongs.
    assert "页 共" not in middle
    assert [hex(ord(c)) for c in middle] == [
        "0x20", "0x9875", "0x3000", "0x5171", "0x20"], \
        [hex(ord(c)) for c in middle]


def test_zh_renders_the_chinese_sentence_end_to_end(tmp_path):
    """The cached (pre-recalculation) text is the sentence a reader sees."""
    _res, xml, _doc = make_footer(tmp_path, template="zh")
    assert repr("".join(texts(xml))) == repr("第 1 页\u3000共 1 页")


# ---------------------------------------------------------------------------
# 2. both fields carry a cached result, so the file shows numbers unopened
# ---------------------------------------------------------------------------
def test_zh_fields_carry_cached_results(tmp_path):
    """``begin``/``instrText``/``separate``/result/``end`` around each field."""
    _res, xml, _doc = make_footer(tmp_path, template="zh")
    frag = body(xml)
    assert frag.count('<w:fldChar w:fldCharType="begin"/>') == 2
    assert frag.count('<w:fldChar w:fldCharType="separate"/>') == 2
    assert frag.count('<w:fldChar w:fldCharType="end"/>') == 2
    # The two result runs hold "1" (cached PAGE, cached total) -- a field with
    # no result shows nothing until somebody presses F9.
    results = [t for t in texts(xml) if t == "1"]
    assert len(results) == 2, results


def test_zh_cached_page_follows_fmt_and_start(tmp_path):
    """``fmt``/``start`` land on the section, and the cache matches."""
    res, xml, _doc = make_footer(tmp_path, template="zh", fmt="lowerRoman",
                                 start=3)
    assert res["cached"] == "iii"
    assert res["fmt"] == "lowerRoman"
    assert res["start"] == "3"
    assert "第 iii 页" in "".join(texts(xml))


# ---------------------------------------------------------------------------
# 3. ``total: sectionpages`` means SECTIONPAGES, zh included
# ---------------------------------------------------------------------------
def test_zh_total_sectionpages_uses_sectionpages_field(tmp_path):
    res, xml, _doc = make_footer(tmp_path, template="zh",
                                 total="sectionpages")
    assert res["total_field"] == "SECTIONPAGES"
    assert instrs(xml) == ["PAGE", "SECTIONPAGES"]
    assert "NUMPAGES" not in body(xml)


def test_zh_default_total_is_numpages(tmp_path):
    res, xml, _doc = make_footer(tmp_path, template="zh")
    assert res["total_field"] == "NUMPAGES"
    assert instrs(xml) == ["PAGE", "NUMPAGES"]


# ---------------------------------------------------------------------------
# 4. backward compatibility: the four old templates are untouched
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("kwargs,expected_instr,expected_texts", [
    ({"template": "page"}, ["PAGE"], []),
    ({"template": "page_of"}, ["PAGE", "NUMPAGES"], ["Page ", " of "]),
    ({"template": "page_of", "prefix": "Seite ", "suffix": " von "},
     ["PAGE", "NUMPAGES"], ["Seite ", " von "]),
    ({"template": "dash"}, ["PAGE"], ["- ", " -"]),
    ({"template": "custom", "prefix": "Kap. ", "suffix": ""},
     ["PAGE"], ["Kap. "]),
    ({}, ["PAGE", "NUMPAGES"], ["Page ", " of "]),
    ({"template": "page_of", "total": "sectionpages"},
     ["PAGE", "SECTIONPAGES"], ["Page ", " of "]),
])
def test_legacy_templates_keep_their_exact_layout(tmp_path, kwargs,
                                                 expected_instr,
                                                 expected_texts):
    """Every pre-existing template still emits the same segments.

    ``custom`` without a placeholder must keep the implicit mid-PAGE: the
    placeholder version moves the fields, and a spec that never asked for one
    must not lose its number.
    """
    _res, xml, _doc = make_footer(tmp_path, **kwargs)
    assert instrs(xml) == expected_instr
    assert text_runs(xml) == expected_texts


def test_custom_without_placeholder_is_byte_identical_to_the_old_layout(
        tmp_path):
    """The pre-change XML, reproduced from the shape rather than re-derived.

    The old implementation emitted exactly ``[text prefix][PAGE][text suffix]``
    with three parts even when a half was empty (the empty part was skipped,
    so the part *count* could differ).  Pinning the whole paragraph here is
    what makes "no behaviour change" a test rather than a claim.
    """
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="Kap. ", suffix=" /")
    assert body(xml) == (
        '<w:p><w:pPr><w:pStyle w:val="Footer"/><w:jc w:val="center"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">Kap. </w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t xml:space="preserve">1</w:t></w:r>'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
        '<w:r><w:t xml:space="preserve"> /</w:t></w:r></w:p>')


# ---------------------------------------------------------------------------
# 5. custom's {page} / {total}
# ---------------------------------------------------------------------------
def test_custom_placeholders_produce_both_fields(tmp_path):
    """``--prefix "{page}/" --suffix " of {total}"`` -> the requested layout."""
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="{page}/", suffix=" of {total}")
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    assert text_runs(xml) == ["/", " of "]
    # No phantom third PAGE: naming {page} is what places the counter.
    assert body(xml).count("<w:fldChar w:fldCharType=\"begin\"/>") == 2
    assert "".join(texts(xml)) == "1/ of 1"


def test_custom_placeholders_in_the_prefix_alone(tmp_path):
    """Both tokens may live in one half; the implicit PAGE must not reappear."""
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="p{page} / {total}")
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    assert text_runs(xml) == ["p", " / "]
    assert "".join(texts(xml)) == "p1 / 1"


def test_custom_placeholders_with_text_on_both_sides(tmp_path):
    """Text before *and* after each placeholder survives, in order."""
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="[p{page}]", suffix="({total})")
    assert text_runs(xml) == ["[p", "]", "(", ")"]
    assert "".join(texts(xml)) == "[p1](1)"


def test_custom_placeholder_total_honours_sectionpages(tmp_path):
    res, xml, _doc = make_footer(tmp_path, template="custom",
                                 prefix="{page}", suffix="/{total}",
                                 total="sectionpages")
    assert res["total_field"] == "SECTIONPAGES"
    assert instrs(xml) == ["PAGE", "SECTIONPAGES"]


def test_custom_stray_braces_are_kept_verbatim(tmp_path):
    """Only the two exact tokens expand; ``{foo}`` is a footer label, not a bug."""
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="{foo}{page}")
    assert text_runs(xml) == ["{foo}"]
    assert instrs(xml) == ["PAGE"]


def test_custom_page_only_placeholder_still_has_no_total(tmp_path):
    """``{page}`` alone must not silently add a NUMPAGES the caller never asked."""
    _res, xml, _doc = make_footer(tmp_path, template="custom",
                                  prefix="Seite {page}.")
    assert instrs(xml) == ["PAGE"]
    assert text_runs(xml) == ["Seite ", "."]


# ---------------------------------------------------------------------------
# 6. zh is fixed: a prefix/suffix is ignored *and reported*
# ---------------------------------------------------------------------------
def test_zh_ignores_prefix_and_suffix_but_says_so(tmp_path):
    res, xml, _doc = make_footer(tmp_path, template="zh", prefix="XX",
                                 suffix="YY")
    assert res["ignored"] == ["prefix", "suffix"]
    assert text_runs(xml) == ["第 ", " 页" + FULLWIDTH_SPACE + "共 ", " 页"]
    assert "XX" not in body(xml) and "YY" not in body(xml)


def test_zh_reports_only_the_half_that_was_supplied(tmp_path):
    res, _xml, _doc = make_footer(tmp_path, template="zh", suffix="YY")
    assert res["ignored"] == ["suffix"]


def test_zh_without_prefix_has_no_ignored_key(tmp_path):
    """The key appears only when something really was dropped."""
    res, _xml, _doc = make_footer(tmp_path, template="zh")
    assert "ignored" not in res


# ---------------------------------------------------------------------------
# 7. run properties ride the new templates too
# ---------------------------------------------------------------------------
def test_zh_text_runs_carry_all_four_rfonts_attributes(tmp_path):
    """The CJK text runs need ``eastAsia`` bound, not just ascii/hAnsi.

    ``run.font.name`` writes two of the four attributes; on ``第 X 页`` that
    leaves the Chinese glyphs on the theme font while the digits use the
    requested one -- a mixed-script jump inside a single footer.
    """
    _res, xml, _doc = make_footer(tmp_path, template="zh",
                                  font="Times New Roman", size_pt=9, bold=True)
    frag = body(xml)
    rfonts = frag.count('<w:rFonts w:ascii="Times New Roman" '
                        'w:hAnsi="Times New Roman" '
                        'w:eastAsia="Times New Roman" w:cs="Times New Roman"/>')
    # three text runs + five runs per field x two fields
    assert rfonts == 3 + 5 * 2, rfonts
    assert frag.count('<w:sz w:val="18"/>') == 3 + 5 * 2


# ---------------------------------------------------------------------------
# 8. the CLI: choices, help, and the failure path
# ---------------------------------------------------------------------------
def test_cli_accepts_template_zh(tmp_path):
    proc, out = run_footer_cli(tmp_path, "--template", "zh")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 0, stderr
    assert "Traceback" not in stderr
    payload = json.loads(proc.stdout.decode("utf-8"))
    assert payload["template"] == "zh"
    assert payload["total_field"] == "NUMPAGES"
    xml = footer_xml(out)
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    assert "".join(texts(xml)) == "第 1 页\u3000共 1 页"


def test_cli_help_lists_zh_among_the_templates(tmp_path):
    proc = subprocess.run(
        [sys.executable, str(SCRIPTS / "docx_section.py"), "pagenums",
         "--help"],
        capture_output=True, check=False,
        env=dict(os.environ, LC_ALL="C", PYTHONIOENCODING="utf-8"))
    help_text = proc.stdout.decode("utf-8", "replace")
    assert proc.returncode == 0
    assert "{page,page_of,zh,dash,custom}" in help_text, help_text
    # --prefix/--suffix must still be documented next to the templates, since
    # their meaning changed for ``custom``.
    assert "--prefix" in help_text and "--suffix" in help_text


def test_cli_rejects_an_unknown_template_with_exit_2(tmp_path):
    """Usage error: argparse refuses before anything is opened or written."""
    proc, out = run_footer_cli(tmp_path, "--template", "zh-CN")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 2, proc.returncode
    assert proc.stdout == b""
    assert "invalid choice" in stderr
    assert not out.exists(), "exit 2 promises nothing was written"


def test_function_rejects_an_unknown_template_and_names_zh():
    """The library's own guard, reached by a caller that bypasses argparse."""
    doc = Document()
    with pytest.raises(SystemExit) as exc:
        ds.add_page_number_field(doc.sections[-1], template="jp")
    message = str(exc.value)
    assert "unknown template" in message
    for name in ("page", "page_of", "zh", "dash", "custom"):
        assert name in message, (name, message)


def test_template_choices_constant_matches_the_branches():
    """One list drives choices + the error message; it must stay honest."""
    assert ds.PAGE_NUM_TEMPLATES == ("page", "page_of", "zh", "dash", "custom")
    assert "zh" in ds.PAGE_NUM_TEMPLATES


# ---------------------------------------------------------------------------
# 9. the spec path (docx_create) reaches the new template
# ---------------------------------------------------------------------------
def spec_with(**fpn):
    settings = {"template": "zh"}
    settings.update(fpn)
    return {"page": {"width_mm": 210, "height_mm": 297,
                     "margins_mm": {"top": 25, "bottom": 25, "left": 26,
                                    "right": 26}},
            "footer_page_numbers": settings,
            "blocks": [{"type": "paragraph", "text": "正文。"}]}


def test_spec_top_level_footer_page_numbers_reaches_zh(tmp_path):
    """``{"footer_page_numbers": {"template": "zh"}}`` is accepted *and* used.

    ``PAGE_NUMBER_KEYS`` validates the *names* of ``add_page_number_field``'s
    keywords, not their values, so nothing there had to change -- this test is
    what keeps that true, because a future value whitelist would break the
    path silently.
    """
    proc, out = run_create_cli(tmp_path, spec_with())
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 0, stderr
    assert "Traceback" not in stderr
    payload = json.loads(proc.stdout.decode("utf-8"))
    assert payload["ok"] is True
    xml = footer_xml(out)
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    assert "".join(texts(xml)) == "第 1 页\u3000共 1 页"


def test_spec_section_break_footer_page_numbers_reaches_zh(tmp_path):
    """The other entry point: a ``section_break`` block's settings."""
    spec = {"page": {"width_mm": 210, "height_mm": 297},
            "blocks": [
                {"type": "paragraph", "text": "封面"},
                {"type": "section_break",
                 "footer_page_numbers": {"template": "zh",
                                         "total": "sectionpages"}},
            ]}
    proc, out = run_create_cli(tmp_path, spec, name="sec.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    doc = Document(out)
    assert len(doc.sections) == 2
    with zipfile.ZipFile(out) as z:
        names = [n for n in z.namelist() if n.startswith("word/footer")]
        xml = z.read(sorted(names)[-1]).decode("utf-8")
    assert instrs(xml) == ["PAGE", "SECTIONPAGES"]
    assert "".join(texts(xml)) == "第 1 页\u3000共 1 页"


def test_spec_custom_placeholders_via_docx_create(tmp_path):
    """``{page}``/``{total}`` survive the spec layer unescaped."""
    proc, out = run_create_cli(
        tmp_path,
        spec_with(template="custom", prefix="{page}/", suffix=" of {total}"),
        name="ph.json")
    assert proc.returncode == 0, proc.stderr.decode("utf-8", "replace")
    xml = footer_xml(out)
    assert instrs(xml) == ["PAGE", "NUMPAGES"]
    assert "".join(texts(xml)) == "1/ of 1"


def test_spec_rejects_a_typo_in_a_footer_key_not_the_template_value(tmp_path):
    """The neighbouring failure path: unknown *key* -> exit 2, one error line."""
    spec = spec_with(templte="zh")
    proc, out = run_create_cli(tmp_path, spec, name="typo.json")
    stderr = proc.stderr.decode("utf-8", "replace")
    assert proc.returncode == 2, proc.returncode
    assert not out.exists()
    assert "unknown key" in stderr
    assert "templte" in stderr
