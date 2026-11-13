# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0] - 2026-11-13

Two halves. The first is P8-P10: an ambiguous invocation is refused instead of
silently doing something else, and none of that changes a document's content.
The second is P12, which does change what a spec produces - headings come out
black, a table can be held on one page, an image can be sized as a share of the
text block, the page-number footer learns Chinese, and a new CLI renders the
result so the layout can be checked without opening Word. Read `Changed` before
re-running an old spec.

### Added

- **`docx_preview.py` - render a document and lint the layout, in one command.**
  A headless LibreOffice produces the PDF; PyMuPDF turns it into per-page PNGs and
  measures the page. Three composable capabilities, at least one required:
  `--pdf OUT.pdf` (default: a temporary file, deleted afterwards), `--pages DIR`
  (one `page-%03d.png` per page, `--dpi` default 110), and `--lint` - per-page
  `fill_ratio`, `sparse_pages` below `--min-fill` (default 0.40), `blank_pages`,
  `image_over_width`, `table_split`, and `orphan_last` with `--max-orphan-chars`
  (default 2). `--json` prints one object, `--text` the default human-readable
  lines; a single run never mixes the two. `--timeout` (default 180) bounds the
  conversion. **`table_split` is a heuristic and may over-report** - it matches
  table geometry across a page boundary, so a table that merely ends near a break
  can be flagged. Needs `soffice` on the path, and PyMuPDF for the PNG and lint
  halves; either one missing is exit `1` with a single `error:` line, never a
  traceback, and `fitz` is imported inside a function so importing the module does
  not require it. The two checks run in the order the work does - convert, then
  read the PDF back - so a machine without both is told about `soffice` first;
  naming PyMuPDF first sent you to install the wrong thing and to hit the same
  wall one step later. This is the file that makes **`scripts/` 27 -> 28** (27 `.py` +
  1 `.ps1`, 26 CLIs of which 19 carry an argparse surface) - every count in
  `SKILL.md`, `SKILL.zh-CN.md`, `README.md` and `README.zh-CN.md` moves with it.
- **`table.keep_together`** - ask Word to hold a whole table on one page. Word has
  no table-level attribute for this, so the request is written in two layers:
  `w:trPr/w:cantSplit` on **every** row (a row may not be split across a page
  break) and `w:keepNext` on every paragraph of every row except the last (each
  row drags the next one along). Default `false`; with the key absent neither
  element is written at all, so an old spec's table XML is untouched. **`cantSplit`
  only forbids splitting a row - a table taller than one page still breaks across
  pages.** `keep_together` is a request, not a guarantee.
- **`image.width_frac`** - size an image as a share of the current section's text
  width, `0 < f <= 1`, resolved through `content_width_mm(section)` so it follows
  the page size and the margins instead of hard-coding a number. A4 with 26 mm
  left and right margins gives a 158 mm text block, so `1.0` is 158.0 mm and `0.5`
  is 79.0 mm. **Mutually exclusive with `width_mm`**: passing both is exit `2` with
  nothing written. Out-of-range values are refused with the range rather than
  clamped, and `0.0` is refused rather than read as "no width" - a number the
  caller wrote is never quietly turned into something else.
- **`docx_section.py pagenums --template zh`** - a Chinese page-number footer,
  `第 <PAGE> 页　共 <NUMPAGES> 页`. The separator between `页` and `共` is a
  full-width space (U+3000), not an ordinary one. `--total sectionpages`
  swaps the second field for `SECTIONPAGES`, so the total counts the section rather
  than the document. `zh` builds its own text and therefore ignores `--prefix` /
  `--suffix`, but reports `"ignored": [...]` rather than dropping them silently.
  The `--template` choices are driven by one constant (`PAGE_NUM_TEMPLATES`), four
  -> five, and the error message reads the same constant.
- **`custom` footer placeholders** - `{page}` and `{total}` inside `--prefix` /
  `--suffix` expand to the PAGE and total fields, split in the order they appear,
  so `--prefix "{page}/" --suffix " of {total}"` gives `1/ of 12`. With no
  placeholder in either string `custom` takes the old code path, and the four old
  templates plus a placeholder-free `custom` are **byte-for-byte unchanged**
  (aggregate SHA-256 identical before and after across nine fixtures). The same
  keys work in a spec, as a top-level `footer_page_numbers` or one inside a
  `section_break`.
- **`tests/test_fix_zoom.py` — 11 tests for the `fix_zoom` contract.** `fix_zoom`
  had no test file at all: `test_table_layout.py` imported its pure
  `fix_settings_xml` and never touched `fix_docx`, so both of that function's
  `return 2` paths were unpinned. The new file covers the refusal paths (a
  non-zip, a missing path, and "refused input must not be rewritten"), the
  success paths (idempotence on an already-valid `w:percent`, repair of a bare
  `<w:zoom/>`), and the passthrough itself — `docx_create.main()` must return `2`
  when `fix_docx` does, and `0` when it does not. The passthrough cases inject a
  stub through `monkeypatch`: the failure branch is not reachable from a real
  build, so an end-to-end test alone would have left the guard unasserted.

### Changed

- **A `styles[]` entry that names an existing style now modifies it in place
  instead of failing.** `{"name": "Heading 1", "size_pt": 16}` used to raise on
  the assumption that a name collision was a mistake, but the built-in `Normal`,
  `Heading 1`-`Heading 9`, `Title` and `Subtitle` are exactly the styles a caller
  most often wants to adjust. Only the keys the entry gives are written - `font`,
  `bold`, `italic`, `base`, `outlineLvl` and the theme-font binding stay as they
  were - so a size-only entry cannot silently undo a colour. A second apply of the
  same entry still leaves one `w:style` behind. This removes one exit-`2` path on
  purpose: a duplicate name is a request now, not an error. **To change a colour
  back, give `color` explicitly.**
- **`Heading 1`-`Heading 9`, `Title` and `Subtitle` are black by default.** A
  neutralising pass runs before any `styles[]` entry: it strips `w:themeColor` /
  `w:themeShade` / `w:themeTint` and writes `w:val="000000"`; the character
  variants (`Heading1Char` ... `TitleChar`, `SubtitleChar`) get the same
  treatment; `Title`'s paragraph bottom border switches to black with `w:sz` and
  `w:space` untouched; and `Subtitle` loses its italic (`w:i` and `w:iCs` -
  `Heading 9` keeps its own). Replacing the `w:color` element is the point: a
  surviving `themeColor` outranks `w:val` in `CT_Color`, so merely writing
  `w:val="000000"` next to it renders theme blue while the XML looks changed. A
  `styles[]` entry that gives `color` still wins, because the pass runs first.
- **`table_caption` followed immediately by a `table` is linked by default**, with
  no key to set: the caption paragraph gets `keep_with_next` and the table's first
  row gets `cantSplit`, so a one-line caption stops being stranded at the foot of
  the previous page. Anything between them - a paragraph, a page break, another
  table - breaks the linkage, because "right after" is measured on the block list.
  **This is a layout change for spec files written before it existed: re-running an
  old spec can paginate differently.** That is the defect being fixed; insert a
  paragraph between the caption and the table to restore the old behaviour.
- **`margins_mm` errors now carry a shape example.** A `margins_mm` passed as a
  list, or an unknown key inside it, gets `; expected shape: {"top": 25, "bottom":
  25, "left": 20, "right": 20}` appended **after** the original sentence, and a
  non-object where an object belongs reports `expected e.g. {"top": .., "bottom":
  .., "left": .., "right": ..}`. Prose only - the validation logic is unchanged and
  the existing `valid keys here` list is preserved verbatim.
- **Usage errors exit `2` with one readable line on stderr and never a traceback.**
  That now covers a negative or out-of-range `--index` (`docx_crossref.py`, and
  `docx_edit.py insert`), an input path that is not a readable `.docx`, an unknown
  `--bookmark-span` / `--vert-align`, and a malformed JSON spec. A negative
  `--index` used to be read as "counted from the end", so a caption or bookmark
  landed on the wrong paragraph and still reported success. Document-state
  failures — an unknown bookmark, a bookmark name already taken — stay on `1`,
  and `3` remains `docx_mathify.py`'s partial-success code alone.
- **`table_width.py --widths` refuses more values than the table has columns** —
  exit `2`, nothing written — instead of dropping the extras on the floor. Fewer
  values still pad with the last one, but the pad is now announced on stderr
  (`warning: table 1: --widths has 1 values but the table has 2 columns; padded 1
  column(s) with 3.00 cm`). `report` on a genuinely table-free document still
  exits `0`; a *mutating* action on one exits `1` with `no tables found; nothing
  written`, because not one width was written.
- **`docx_toc.py --pages` must carry exactly one value per entry** at the levels
  selected by `--levels`; a different count exits `2` and writes no file. Surplus
  values used to be dropped and missing ones left as blank slots, both silently,
  and a TOC with a blank page number in the middle reads exactly like one that was
  never paginated. The Python API stays tolerant on purpose — a caller filling
  only some slots is a different thing from a mistyped comma.
- **An unknown key in a `docx_create` spec is rejected before anything is
  written**: exit `2`, `{"ok": false, …}` on stdout, and one stderr line naming the
  offending key plus the nearest legitimate spelling. A misspelt key used to build
  a document with the content missing and exit `0`. Nested objects are checked too
  — block fields, `runs[]`, `tab_stops[]`, `page`, `page.margins_mm`, `styles[]`,
  `footer_page_numbers` — and out-of-range values are reported with their real
  range instead of being clamped.
- **`docx_mathify.py --inplace` together with `-o` is a usage error** — exit `2`,
  both files untouched, no `.bak`. `--inplace` used to win silently: the `-o` path
  was never created and the *input* file was overwritten instead. A run with none
  of the three landing flags still exits `2`. `insert_formula.py` follows the same
  rule for `--dump` with
  `--after`/`--before` and for an empty marker, and now prints the anchor it
  measured (`… at body[1]`, `at tc[i]` inside a table cell) rather than echoing
  the argument.
- **`docx_template.py --fix-formulas` and `--no-fix-formulas` are mutually
  exclusive** (argparse exit `2`). Both used to be accepted and took effect as an
  AND: the default marker pass switched off while the OMML pass received raw
  `_xx_` markers, the conversion failed, and the CLI still exited `0`.
- **`docx_edit.py page-numbers` is idempotent.** A footer that already carries a
  `PAGE`/`NUMPAGES` field — either the complex `fldChar` form or the
  `w:fldSimple` one — is reported `{"changed": false, …}`, and an in-place re-run
  writes nothing at all, so plan → apply → apply no longer staples a second
  "Page 1 of 1" onto the first.
- **`docx_pack.unpack` refuses two more entry shapes.** Windows reserved device
  names (`NUL`, `CON`, `PRN`, `AUX`, `COM0`-`COM9`, `LPT0`-`LPT9`, `CONIN$`,
  `CONOUT$`, extension ignored) and a last component with a trailing dot or space.
  `NUL` accepted the write and discarded it, so `unpack` reported success for a
  part that never reached the disk; `evil.txt ` is created as `evil.txt`, so two
  distinct entries collide on one file.

### Fixed

- **`docx_create.py` writes `error:` on stderr where it used to write
  `docx_create:`.** Every other CLI already used the `error:` prefix that D-024
  makes the single shape of a refusal line, so a caller branching on the first
  token treated this one script's refusals as a different class of message.
  Nothing in the repository read the old prefix. The stdout JSON contract is
  untouched: `docx_create` is one of the scripts documented as emitting JSON on
  stdout, and a refusal still writes `{"ok": false, ...}` there - the two channels
  are deliberately not both empty.
- **`docx_mathify.py` refuses every pair of landing flags that disagree with each
  other.** `--inplace` together with `--dry-run` is an argparse usage error (exit
  `2`), and so is `-o` together with `--dry-run` (exit `2`, one stderr line, no
  file created). `--dry-run` used to win over `--inplace` silently, and `-o` with
  `--dry-run` used to print `输出: <path>` for a file that was never written --
  a reported output path that no pass would ever produce is the same defect as an
  echoed `--index`, just in a different flag. No silent landing-flag combination
  remains.
- **Page-number footers with `font`, `size_pt` or `bold` no longer crash.**
  `docx_section._rpr` built its `w:rFonts` element from bare `qn("ascii")` /
  `OxmlElement("b")` names, so every one of those three options raised
  `KeyError`/`ValueError` and `footer_page_numbers` was unusable in the shape the
  documentation described. The qualified names are now used in all three
  branches. Note that this helper only decorates the literal text run, not the
  `PAGE` field result: `bold: false` and omitting `bold` are not equivalent, and
  neither affects the number Word renders.
- **`docx_toc.py --insert-at` and `--section-index` refuse an index that names no
  real position** — exit `2`, nothing written. `--insert-at 999` on a short body
  used to append and report `inserted_at: null` with exit `0`; `--insert-at -1`
  used to hit Python's negative indexing and land before the *last* paragraph
  while echoing the argument back. A rebuild checks the index twice, before and
  after the previous TOC is removed, because that removal shortens the body. An
  out-of-range `--section-index` used to fall back to section 0 silently, which
  put the tab stops on the wrong section's geometry — 40 mm off on the sample;
  it now exits `2`. Negative `--section-index` (`-1` = last section) is kept.
- **`fix_zoom.py` reports an unreadable input as a usage error** (`2`) rather
  than a traceback, and `docx_create.py` now propagates that return code instead
  of discarding it. `fix_docx()` is called right after `doc.save()` to patch the
  schema-required `w:percent` that python-docx omits; its result used to be
  dropped, so the one failure it can report — a package that is not a readable
  zip — reached the caller as `{"ok": true}` and exit `0`. That path is narrow in
  practice (the file was written a line earlier), and the guard is what makes it
  closed rather than merely unlikely: any future failure mode `fix_docx()` learns
  to report will surface without another edit here.

## [1.0.0] - 2026-11-11

First tagged release. The toolkit itself had been published earlier on `main`;
this tag marks the point where the three pre-release blockers below were
closed and the sample data in the tests was replaced with neutral fixtures.

### Fixed

- **`docx_pack.unpack` accepted Windows drive prefixes.** A zip entry named
  `D:/x.txt` was extracted to the drive root: `_safe_member` only rejected names
  beginning with `/`, and `pathlib` discards the left operand when joining an
  absolute path. Relative traversal (`../`) and POSIX absolute paths were
  already blocked, so only the Windows form escaped. `_safe_member` now rejects
  any name for which `ntpath.splitdrive` yields a non-empty drive, which also
  covers `D:x` and UNC paths; on POSIX `splitdrive` always returns an empty
  drive, so the check is inert there.
- **`docx_template.py apply --to X -o Y` ignored `-o`.** The `--to` branch never
  read `args.output`, so the documented command rewrote `X` in place, produced
  no `Y`, left no backup, and still exited 0. With `-o` the template is now
  applied to a copy and `X` is never written; without `-o` the in-place write
  leaves a `.bak` first. Pointing `-o` at a missing directory is now an error
  instead of a silent overwrite.
- **`docx_pack.py` exited 0 for any argument while doing nothing.** A caller who
  mistook the library module for a CLI saw a silent success, which then made the
  next command look like a failure of the document rather than of the invocation.
  `--help` and no arguments still exit 0 with the API summary; any other argument
  reports the mistake on stderr and exits 2.

### Added

- 71 tests across four new or extended files (`pytest <files> --collect-only`:
  38 + 14 + 7 + 12):
  - `tests/test_docx_edit_subcommands.py` — 38, all eight `docx_edit` subcommands,
    including cross-run `replace` and a `normalize` case that asserts the merged
    runs keep every character.
  - `tests/test_docx_pack_safety.py` — 14, drive-prefix, POSIX-absolute and
    relative traversal, each asserting both the refusal and that nothing was
    written.
  - `tests/test_docx_template_output.py` — 7, `-o` honoured, source untouched,
    `.bak` on in-place writes.
  - `tests/test_insert_formula.py` — 12, the `--display`, `--dump` and
    `--raw-spacing` switches and a forward `--before` case.
- `## Requirements` in the English README, and a platform-boundary note in both
  languages: the formula chain needs `MML2OMML.XSL` from a local Microsoft
  Office, `find_xslt()` only probes Windows roots, and PDF export drives Word
  over COM.

### Changed

- The CLI/library split is documented consistently as 25 command-line scripts
  and 1 pure library (`_bootstrap.py`). Earlier revisions of the docs disagreed
  with each other and with the code.
- CI floors move to the newly measured runner numbers: `passed >= 1356`,
  `skipped <= 372`. The floor only ever moves up, so losing a block of tests
  trips it rather than passing silently.
- The sample data used by the regression tests no longer carries the identifiers
  of a real submission. Structure is unchanged — the 27-entry outline and its
  levels are intact — only the wording is neutral now.

### Notes

- `--fix-formulas` runs a second in-place pass after the template merge; the
  `.bak` snapshot predates both passes, so there is no separate rollback point
  between them.
- An existing `X.bak` is overwritten on the next in-place apply. The command
  says so on stdout; it does not rotate backups.

[1.1.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.1.0
[1.0.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.0.0