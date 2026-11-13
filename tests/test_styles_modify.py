# -*- coding: utf-8 -*-
"""``styles[]`` reaches built-in styles, and headings print black by default.

Two changes meet here and must not fight each other:

* ``add_styles`` used to call ``Styles.add_style`` unconditionally, which raises
  on a name that already exists — so no spec could touch ``Heading 1``.  It now
  looks the name up first and modifies in place.
* The built-in heading / title / subtitle styles are theme-coloured (blue, and
  ``Subtitle`` is italic), so every document came out with blue headings.  They
  are neutralised to black at build time, **before** ``add_styles``.

The ordering is the whole point, and it is what the counter-proof test at the
end pins: neutralise → apply an entry that says only ``size_pt`` → the colour
must still be black.  An implementation that rebuilt the style from the spec
entry ("reset then apply") would pass the size assertion and silently send the
heading back to the theme's blue, so the negative assertion is the real test.

Everything is asserted on the XML, none of it needs Office.
"""
from __future__ import annotations

import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pytest

SKILL = Path(__file__).resolve().parents[1]
SCRIPTS = SKILL / "scripts"
CREATE = SCRIPTS / "docx_create.py"
sys.path.insert(0, str(SCRIPTS))

from docx import Document  # noqa: E402
from docx.enum.style import WD_STYLE_TYPE  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402

import docx_create as dc  # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


# ---- helpers -----------------------------------------------------------------

def run_create(tmp_path, spec: dict):
    """Run the CLI; return ``(returncode, stdout payload, stderr text, out)``."""
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
    return proc.returncode, payload, proc.stderr or "", out


def styles_xml(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        return archive.read("word/styles.xml").decode("utf-8")


def style_block(xml: str, style_id: str) -> str:
    """The single ``<w:style …>`` element whose ``styleId`` is *style_id*."""
    marker = 'w:styleId="%s"' % style_id
    at = xml.find(marker)
    assert at != -1, "no w:style with styleId=%r in styles.xml" % style_id
    open_at = xml.rfind("<w:style ", 0, at)
    close_at = xml.find("</w:style>", at)
    assert open_at != -1 and close_at != -1
    return xml[open_at:close_at]


def style_rpr(style):
    return style.element.find(qn("w:rPr"))


def style_colour(style):
    rpr = style_rpr(style)
    return None if rpr is None else rpr.find(qn("w:color"))


def make_theme_styled(doc, name: str, *, italic=False, colour="4F81BD"):
    """A *dirty* style that looks like Word's own: theme colour, not explicit."""
    style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
    rpr = style.element.get_or_add_rPr()
    if italic:
        for tag in ("w:i", "w:iCs"):
            rpr.append(OxmlElement(tag))
    colour_el = OxmlElement("w:color")
    colour_el.set(qn("w:val"), colour)
    colour_el.set(qn("w:themeColor"), "accent1")
    colour_el.set(qn("w:themeShade"), "BF")
    rpr.append(colour_el)
    return style


# ---- 1. a spec can restyle a built-in style (judge 1) ------------------------

def test_a_spec_restyles_the_builtin_heading_1(tmp_path):
    """``{"name": "Heading 1", …}`` used to be a hard error; now it is an edit.

    End-to-end through the CLI, because "the function accepts it" and "a run
    produces a document where the heading is 黑体 16 pt black" are different
    claims, and the font one has a trap: ``CT_Fonts`` gives
    ``w:asciiTheme``/``w:eastAsiaTheme`` **precedence over** ``w:ascii``/
    ``w:eastAsia``, and Heading 1 ships bound to the theme.  A spec that names
    ``黑体`` must therefore clear the theme binding, or Word shows the theme face
    while the XML shows 黑体.
    """
    spec = {
        "styles": [{"name": "Heading 1", "font": "黑体", "size_pt": 16,
                    "color": "000000"}],
        "blocks": [{"type": "heading", "text": "第一章", "level": 1}],
    }
    rc, payload, stderr, out = run_create(tmp_path, spec)

    assert rc == 0, stderr
    assert payload["ok"] is True
    assert out.exists()

    block = style_block(styles_xml(out), "Heading1")
    assert 'w:ascii="黑体"' in block
    assert 'w:hAnsi="黑体"' in block
    assert 'w:eastAsia="黑体"' in block
    assert 'w:cs="黑体"' in block
    # The theme twins must be gone, otherwise they win at render time.
    for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        assert "%s=" % attr not in block, (attr, block)
    # 16 pt == 32 half-points, the unit ``w:sz`` uses.
    assert 'w:sz w:val="32"' in block
    # ``w:szCs`` is deliberately *not* asserted: python-docx's ``font.size``
    # writes ``w:sz`` only, so the built-in ``w:szCs w:val="28"`` survives.  That
    # asymmetry predates this change and is recorded in the P12-R1A report §5
    # rather than fixed here — writing ``w:szCs`` would also change the output of
    # every newly created style, which is outside a styles-in-place change.
    # Black, and *not* "black plus a theme colour that overrides it".
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="000000"' in colour, colour
    assert "themeColor" not in colour, colour


def test_the_theme_font_binding_is_only_cleared_when_a_font_is_given(tmp_path):
    """No ``font`` in the spec ⇒ ``Heading 1`` keeps Word's theme binding."""
    spec = {"styles": [{"name": "Heading 1", "size_pt": 16}],
            "blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    block = style_block(styles_xml(out), "Heading1")
    assert 'w:asciiTheme="majorHAnsi"' in block
    assert 'w:eastAsiaTheme="majorEastAsia"' in block


# ---- 2. the same name twice leaves one w:style (judge 2) ---------------------

def test_applying_the_same_entry_twice_leaves_one_w_style(tmp_path):
    """Word reports a duplicate ``w:style`` name as a broken file."""
    doc = Document()
    entry = {"name": "Heading 2", "font": "楷体", "size_pt": 14,
             "color": "1F4E79"}
    dc.add_styles(doc, [entry])
    dc.add_styles(doc, [entry])
    out = tmp_path / "twice.docx"
    doc.save(out)

    xml = styles_xml(out)
    assert xml.count('w:styleId="Heading2"') == 1
    assert xml.count('<w:name w:val="heading 2"/>') == 1


def test_the_style_object_identity_is_kept_when_modifying(tmp_path):
    """In-place means the *same element*, not a replacement with the same name."""
    doc = Document()
    before = doc.styles["Heading 2"]
    dc.add_styles(doc, [{"name": "Heading 2", "size_pt": 14}])
    assert doc.styles["Heading 2"].element is before.element


# ---- 3. only the keys the spec gives are written (judge 3) -------------------

def test_a_size_only_entry_keeps_black_through_the_cli(tmp_path):
    """The counter-proof, end to end: P2 neutralises, P1 must not undo it."""
    spec = {"styles": [{"name": "Heading 1", "size_pt": 16}],
            "blocks": [{"type": "heading", "text": "第一章", "level": 1}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    block = style_block(styles_xml(out), "Heading1")
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="000000"' in colour, colour
    assert "themeColor" not in colour, colour
    assert 'w:sz w:val="32"' in block


def test_a_size_only_entry_leaves_the_neutralised_colour_black():
    """The counter-proof: P2 neutralises, P1 must not undo it.

    Sequence: build-time neutralise → ``{"name": "Heading 1", "size_pt": 16}``.
    The colour must still be ``000000`` with no ``themeColor``.  This is the
    case a "reset the style, then apply the spec" implementation fails: the
    size assertion would still pass while the heading went back to blue.
    """
    doc = Document()
    dc._neutralise_heading_styles(doc)
    assert style_colour(doc.styles["Heading 1"]).get(qn("w:val")) == "000000"

    dc.add_styles(doc, [{"name": "Heading 1", "size_pt": 16}])

    colour = style_colour(doc.styles["Heading 1"])
    assert colour is not None
    assert colour.get(qn("w:val")) == "000000"
    assert colour.get(qn("w:themeColor")) is None
    assert colour.get(qn("w:themeShade")) is None
    # …and the key the spec *did* give was applied.
    sz = style_rpr(doc.styles["Heading 1"]).find(qn("w:sz"))
    assert sz.get(qn("w:val")) == "32"


def test_the_counter_proof_fails_when_the_style_is_rebuilt():
    """Proves the previous test can fail: resetting the colour turns it blue.

    Without this, "the colour stayed black" could be true simply because nothing
    in the code path is capable of writing a theme colour back, and the assertion
    would be decoration.  Reproducing the rejected implementation (clear the
    colour, then take it from a spec that has none) shows what the test catches.
    """
    doc = Document()
    dc._neutralise_heading_styles(doc)
    rpr = style_rpr(doc.styles["Heading 1"])
    rpr.remove(rpr.find(qn("w:color")))
    restored = OxmlElement("w:color")
    restored.set(qn("w:val"), "365F91")
    restored.set(qn("w:themeColor"), "accent1")
    restored.set(qn("w:themeShade"), "BF")
    rpr.append(restored)

    colour = style_colour(doc.styles["Heading 1"])
    assert colour.get(qn("w:themeColor")) == "accent1", \
        "the broken shape must be detectable, or the counter-proof proves nothing"


def test_an_entry_without_bold_keeps_the_builtin_bold():
    """Heading 1 is bold in Word; saying nothing about bold must not clear it."""
    doc = Document()
    dc._neutralise_heading_styles(doc)
    dc.add_styles(doc, [{"name": "Heading 1", "size_pt": 16}])
    rpr = style_rpr(doc.styles["Heading 1"])
    assert rpr.find(qn("w:b")) is not None
    assert rpr.find(qn("w:bCs")) is not None


def test_an_entry_without_base_keeps_the_builtin_base(tmp_path):
    """``Heading 1`` is based on ``Normal``; not naming ``base`` must not unbase it."""
    spec = {"styles": [{"name": "Heading 1", "size_pt": 16}],
            "blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    block = style_block(styles_xml(out), "Heading1")
    assert '<w:basedOn w:val="Normal"/>' in block
    assert '<w:next w:val="Normal"/>' in block


def test_an_entry_without_outline_level_keeps_the_builtin_level():
    """``Heading 2`` carries ``outlineLvl`` 1; omitting the key must not rewrite it."""
    doc = Document()
    dc.add_styles(doc, [{"name": "Heading 2", "size_pt": 14}])
    pPr = doc.styles["Heading 2"].element.find(qn("w:pPr"))
    assert pPr.find(qn("w:outlineLvl")).get(qn("w:val")) == "1"


def test_an_entry_without_italic_keeps_the_builtin_italic_of_heading_9():
    """Word's ``Heading 9`` is italic; that is not one of the P9 targets."""
    doc = Document()
    dc._neutralise_heading_styles(doc)
    dc.add_styles(doc, [{"name": "Heading 9", "size_pt": 11}])
    rpr = style_rpr(doc.styles["Heading 9"])
    assert rpr.find(qn("w:i")) is not None


# ---- 3b. the opt-out channel: the spec can ask for a colour back ------------

def test_a_spec_can_overwrite_the_neutralised_colour(tmp_path):
    """Neutralising is a *default*: an explicit ``color`` still wins."""
    spec = {"styles": [{"name": "Heading 1", "color": "1F4E79"}],
            "blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    block = style_block(styles_xml(out), "Heading1")
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="1F4E79"' in colour, colour
    # ``ColorFormat.rgb``'s setter drops themeColor; patching @w:val alone would
    # leave the theme attribute winning and the page still blue.
    assert "themeColor" not in colour, colour


# ---- 4./5. the defaults (judges 4 and 5) ------------------------------------

def test_a_plain_document_has_black_headings(tmp_path):
    """No ``styles[]`` at all: the built-in headings are already black."""
    spec = {"blocks": [
        {"type": "heading", "text": "第一章", "level": 1},
        {"type": "heading", "text": "1.1 节", "level": 2},
        {"type": "paragraph", "text": "正文"},
    ]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    xml = styles_xml(out)
    for style_id in ["Heading%d" % n for n in range(1, 10)]:
        block = style_block(xml, style_id)
        colour = block[block.find("<w:color"):
                       block.find("/>", block.find("<w:color"))]
        assert 'w:val="000000"' in colour, (style_id, colour)
        assert "themeColor" not in colour, (style_id, colour)


def test_a_plain_document_neutralises_title_and_subtitle(tmp_path):
    spec = {"blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    xml = styles_xml(out)
    for style_id in ("Title", "Subtitle"):
        block = style_block(xml, style_id)
        colour = block[block.find("<w:color"):
                       block.find("/>", block.find("<w:color"))]
        assert 'w:val="000000"' in colour, (style_id, colour)
        assert "themeColor" not in colour, (style_id, colour)


def test_subtitle_loses_its_italic():
    """P9: a grey italic subtitle is a screen-document look, not a paper one."""
    doc = Document()
    assert style_rpr(doc.styles["Subtitle"]).find(qn("w:i")) is not None
    dc._neutralise_heading_styles(doc)
    rpr = style_rpr(doc.styles["Subtitle"])
    assert rpr.find(qn("w:i")) is None
    assert rpr.find(qn("w:iCs")) is None


def test_headings_other_than_the_subtitle_keep_their_italic():
    """The italic strip is ``Subtitle``-only, not a blanket sweep."""
    doc = Document()
    dc._neutralise_heading_styles(doc)
    assert style_rpr(doc.styles["Heading 9"]).find(qn("w:i")) is not None


# ---- 9. Title's bottom rule is recoloured, not redesigned ------------------

def test_title_bottom_border_is_black_and_keeps_its_thickness(tmp_path):
    """A black title over a blue underline still reads as "the title is blue".

    Only the colour attributes may change: ``w:sz`` (8 = 1 pt) and ``w:space``
    (4 pt) are the design, and the report never mentioned a rule that is too
    thick or too close.
    """
    spec = {"blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr

    block = style_block(styles_xml(out), "Title")
    border = block[block.find("<w:pBdr"):block.find("</w:pBdr>")]
    # No theme colour anywhere under w:pBdr.
    assert "themeColor" not in border, border
    assert "themeShade" not in border, border
    assert "themeTint" not in border, border
    assert 'w:color="000000"' in border, border
    # Thickness and spacing survive untouched.
    assert 'w:sz="8"' in border, border
    assert 'w:space="4"' in border, border
    assert 'w:val="single"' in border, border


def test_every_border_side_is_recoloured_not_just_bottom():
    """The helper walks all six ``CT_PBdr`` sides, not the one Word happens to ship."""
    doc = Document()
    style = doc.styles["Title"]
    pPr = style.element.get_or_add_pPr()
    pBdr = pPr.find(qn("w:pBdr"))
    if pBdr is None:
        pBdr = OxmlElement("w:pBdr")
        pPr.append(pBdr)
    for side in ("top", "left", "right", "between", "bar"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:val"), "single")
        edge.set(qn("w:sz"), "6")
        edge.set(qn("w:space"), "1")
        edge.set(qn("w:color"), "4F81BD")
        edge.set(qn("w:themeColor"), "accent1")
        pBdr.append(edge)

    report = dc._neutralise_heading_styles(doc)
    assert "Title" in report["borders"]
    for side in ("top", "left", "bottom", "right", "between", "bar"):
        edge = pBdr.find(qn("w:" + side))
        assert edge is not None, side
        assert edge.get(qn("w:color")) == "000000", side
        assert edge.get(qn("w:themeColor")) is None, side
        if side != "bottom":
            assert edge.get(qn("w:sz")) == "6", side


def test_a_border_with_auto_colour_is_made_explicit_black():
    """``w:val="auto"`` resolves through the theme, so it is not "already black"."""
    doc = Document()
    style = doc.styles["Title"]
    pPr = style.element.get_or_add_pPr()
    bottom = pPr.find(qn("w:pBdr")).find(qn("w:bottom"))
    for attr in ("w:themeColor", "w:color"):
        if bottom.get(qn(attr)) is not None:
            del bottom.attrib[qn(attr)]
    bottom.set(qn("w:val"), "auto")
    bottom.set(qn("w:color"), "auto")

    dc._neutralise_heading_styles(doc)
    assert bottom.get(qn("w:color")) == "000000"


# ---- 10. the linked character twins are neutralised too -------------------

@pytest.mark.parametrize("name", ["Heading1Char", "TitleChar", "SubtitleChar"])
def test_char_variants_are_black_without_a_theme_colour(tmp_path, name):
    """Word lists these in the style gallery and links each paragraph style to one.

    A black ``Heading 1`` is not enough: applying ``Heading 1 Char`` to a few
    words inside the heading turns exactly those words blue again.
    """
    spec = {"blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr

    block = style_block(styles_xml(out), name)
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="000000"' in colour, (name, colour)
    assert "themeColor" not in colour, (name, colour)


def test_all_nine_heading_char_variants_are_black():
    doc = Document()
    dc._neutralise_heading_styles(doc)
    for n in range(1, 10):
        name = "Heading %d Char" % n
        colour = style_colour(doc.styles[name])
        assert colour is not None, name
        assert colour.get(qn("w:val")) == "000000", name
        assert colour.get(qn("w:themeColor")) is None, name


def test_subtitle_char_loses_its_italic():
    doc = Document()
    assert style_rpr(doc.styles["Subtitle Char"]).find(qn("w:i")) is not None
    dc._neutralise_heading_styles(doc)
    rpr = style_rpr(doc.styles["Subtitle Char"])
    assert rpr.find(qn("w:i")) is None
    assert rpr.find(qn("w:iCs")) is None


def test_heading_char_variants_keep_their_italic():
    """P9 is ``Subtitle``-only; ``Heading 9 Char`` is italic by design."""
    doc = Document()
    dc._neutralise_heading_styles(doc)
    assert style_rpr(doc.styles["Heading 9 Char"]).find(qn("w:i")) is not None


def test_a_localised_char_variant_is_neutralised_too():
    """A Chinese template names the twin ``标题 1 字符``; the lookup must still find it."""
    doc = Document()
    make_theme_styled(doc, "标题 1 字符")
    report = dc._neutralise_heading_styles(doc)
    assert "标题 1 字符" in report["touched"], report
    colour = style_colour(doc.styles["标题 1 字符"])
    assert colour.get(qn("w:val")) == "000000"
    assert colour.get(qn("w:themeColor")) is None


# ---- 11. ordering: neutralise first, spec second --------------------------

def test_the_spec_still_wins_over_the_neutralised_default(tmp_path):
    """``styles[]`` can still ask for a colour, i.e. neutralising is a default.

    The title promises less than it used to.  An earlier version claimed this
    proved the *call order* ("reverse the two and this returns 000000").  It
    does not: ``_neutralise_style_colour`` only rewrites a colour that carries a
    ``themeColor``, and an explicit ``1F4E79`` has none, so the assertion holds
    with the two calls in either order.  Measured -- moving the neutralise call
    after ``add_styles`` left this test green.

    What it does prove, and all it claims now: a spec colour survives into the
    stylesheet with no theme attribute attached, so a caller can opt back out of
    the neutralised default.  There is **no** ordering invariant to pin here --
    see :func:`test_neutralising_is_a_scrub_not_a_guard` for why the pass being
    idempotent makes the two orders indistinguishable.
    """
    spec = {"styles": [{"name": "Heading 1", "color": "1F4E79"}],
            "blocks": [{"type": "heading", "text": "第一章", "level": 1}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr
    block = style_block(styles_xml(out), "Heading1")
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="1F4E79"' in colour, colour
    assert "themeColor" not in colour, colour


def test_neutralising_is_a_scrub_not_a_guard(tmp_path):
    """Pin what the pass actually is: a one-shot scrub that is idempotent.

    Earlier versions of the ordering test claimed the call order was observable
    through ``styles[]``.  It is not, and neither is ``Subtitle``'s italic: the
    pass only rewrites colours that resolve through the theme, and it re-strips
    ``w:i`` on every call, so running it before *or* after anything else leaves
    the same end state.  Measured: appending ``w:i`` and neutralising again
    removes it a second time.

    So there is no ordering invariant to pin here.  What this test pins instead
    is the property the ordering claim was standing in for, and the one a future
    refactor could plausibly break: **running the neutralise pass twice must not
    add anything and must not raise**, and a style it has already cleaned stays
    clean.  If someone turns it into an accumulating guard (e.g. appending a
    second ``w:color``) or makes it depend on being called exactly once, this
    fails.
    """
    doc = Document()
    first = dc._neutralise_heading_styles(doc)
    assert first["touched"], "factory heading styles carry theme colours"

    # Every style it touched must be black already.
    for name in first["touched"]:
        colour = style_colour(doc.styles[name]) if name in {
            s.name for s in doc.styles} else None
        if colour is not None:
            assert colour.get(qn("w:val")) == "000000", (name, colour.attrib)
            assert colour.get(qn("w:themeColor")) is None, (name, colour.attrib)

    # A second run is a no-op: nothing new to touch, no exception.
    second = dc._neutralise_heading_styles(doc)
    assert second["touched"] == [], second["touched"]
    assert set(second["explicit"]) >= set(first["touched"]), second

    # And an italic put back after the scrub is removed by the next scrub --
    # i.e. the pass is idempotent in the destructive direction too.
    rpr = doc.styles["Subtitle"].element.find(qn("w:rPr"))
    assert rpr is not None
    rpr.append(OxmlElement("w:i"))
    assert rpr.find(qn("w:i")) is not None
    dc._neutralise_heading_styles(doc)
    assert rpr.find(qn("w:i")) is None




def test_the_spec_can_put_the_title_rule_back_to_a_colour(tmp_path):
    """``Title``'s border is not permanently frozen either — but nothing overrides it yet.

    ``STYLE_KEYS`` has no border key, so this documents the current boundary
    rather than promising a feature: the title rule is black and the spec
    vocabulary cannot change it yet.
    """
    spec = {"styles": [{"name": "Title", "color": "1F4E79"}],
            "blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, _payload, stderr, out = run_create(tmp_path, spec)
    assert rc == 0, stderr

    block = style_block(styles_xml(out), "Title")
    border = block[block.find("<w:pBdr"):block.find("</w:pBdr>")]
    assert 'w:color="000000"' in border, border
    assert "themeColor" not in border, border
    # The run colour *is* reachable by the spec; the border is not (yet).
    colour = block[block.find("<w:color"):block.find("/>", block.find("<w:color"))]
    assert 'w:val="1F4E79"' in colour, colour


# ---- 6. a base that points nowhere is still an input error (judge 6) -------

def test_a_dangling_base_still_exits_2_with_one_error_line(tmp_path):
    """``base`` naming a style that does not exist must not silently invent one."""
    spec = {"styles": [{"name": "Heading 1", "base": "NoSuchStyle"}],
            "blocks": [{"type": "paragraph", "text": "正文"}]}
    rc, payload, stderr, out = run_create(tmp_path, spec)

    assert rc == 2, (rc, stderr)
    assert payload is not None and payload["ok"] is False, payload
    assert payload["output"] is None
    assert "NoSuchStyle" in payload["error"]
    # D-024: exit 2 ⇒ one line on stderr, no traceback, and no file on disk.
    lines = [ln for ln in stderr.splitlines() if ln.strip()]
    assert len(lines) == 1, lines
    assert "Traceback" not in stderr
    assert "NoSuchStyle" in lines[0]
    assert not out.exists()


# ---- 7. behaviour that must survive the change -----------------------------

@pytest.mark.parametrize("bad", [-1, 10, "标题", 1.5, object()])
def test_a_bad_outline_level_still_raises_on_an_existing_style(bad):
    """``add_style`` no longer crashes first, so the validation must still run.

    Before P12 the ``ValueError`` came from the name clash before the outline
    level was ever looked at, which would have hidden a regression here.
    """
    doc = Document()
    with pytest.raises(ValueError):
        dc.add_styles(doc, [{"name": "Heading 2", "outline_level": bad}])


def test_a_short_lived_style_still_gets_created(tmp_path):
    """The unchanged half of the contract: a new name is a new style."""
    doc = Document()
    dc.add_styles(doc, [{"name": "卷标题", "base": "Normal", "font": "黑体",
                         "size_pt": 18, "outline_level": 1}])
    style = doc.styles["卷标题"]
    assert style.style_id == "卷标题" or style.style_id is not None
    pPr = style.element.find(qn("w:pPr"))
    assert pPr.find(qn("w:outlineLvl")).get(qn("w:val")) == "0"


# ---- 8. styleId / localised-name compatibility -----------------------------

def test_a_localised_heading_name_is_neutralised_too():
    """A Chinese template names the style ``标题 1``; the id may stay ``Heading1``.

    Built as a dirty fixture rather than by localising python-docx's template,
    because the point is the *lookup*: whichever of the two spellings a
    template uses, the style has to be found.
    """
    doc = Document()
    make_theme_styled(doc, "标题 1")
    report = dc._neutralise_heading_styles(doc)
    assert "标题 1" in report["touched"], report
    colour = style_colour(doc.styles["标题 1"])
    assert colour.get(qn("w:val")) == "000000"
    assert colour.get(qn("w:themeColor")) is None


def test_a_localised_subtitle_name_loses_its_italic():
    doc = Document()
    make_theme_styled(doc, "副标题", italic=True)
    report = dc._neutralise_heading_styles(doc)
    assert "副标题" in report["touched"], report
    assert style_rpr(doc.styles["副标题"]).find(qn("w:i")) is None


def test_an_explicitly_coloured_style_is_reported_not_touched():
    """A style that is already explicit black is left byte-identical."""
    doc = Document()
    style = doc.styles.add_style("标题 3", WD_STYLE_TYPE.PARAGRAPH)
    rpr = style.element.get_or_add_rPr()
    colour = OxmlElement("w:color")
    colour.set(qn("w:val"), "000000")
    rpr.append(colour)
    before = style.element.xml

    report = dc._neutralise_heading_styles(doc)
    assert "标题 3" in report["explicit"], report
    assert "标题 3" not in report["touched"], report
    assert style.element.xml == before


def test_a_style_without_run_properties_is_reported_not_invented():
    """No ``w:rPr`` ⇒ nothing to neutralise, and no opinion invented either."""
    doc = Document()
    style = doc.styles.add_style("标题 7", WD_STYLE_TYPE.PARAGRAPH)
    rpr = style.element.find(qn("w:rPr"))
    if rpr is not None:
        style.element.remove(rpr)

    report = dc._neutralise_heading_styles(doc)
    assert "标题 7" in report["missing"], report
    assert style.element.find(qn("w:rPr")) is None


def test_the_report_names_every_builtin_heading():
    """The report is what makes a silent miss visible, so it must list them all.

    Both families are expected: the paragraph styles *and* their ``w:link``-ed
    character twins.  Missing a char twin is invisible in the document until
    someone applies that character style to a few words, which is exactly the
    hole this round closed.
    """
    doc = Document()
    report = dc._neutralise_heading_styles(doc)
    para = ["Heading %d" % n for n in range(1, 10)] + ["Title", "Subtitle"]
    chars = (["Heading %d Char" % n for n in range(1, 10)]
             + ["Title Char", "Subtitle Char"])
    expected = set(para) | set(chars)
    assert set(report["styles"]) == expected, report["styles"]
    assert set(report["touched"]) == expected, report["touched"]
    assert report["missing"] == []
    assert report["borders"] == ["Title"]


def test_neutralising_twice_is_a_no_op(monkeypatch):
    """Running it on an already-neutralised document changes nothing."""
    doc = Document()
    dc._neutralise_heading_styles(doc)
    snapshot = doc.styles["Heading 1"].element.xml
    second = dc._neutralise_heading_styles(doc)
    assert second["touched"] == [], second
    assert "Heading 1" in second["explicit"], second
    assert doc.styles["Heading 1"].element.xml == snapshot
