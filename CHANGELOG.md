# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

Everything below changes what a caller *sees* on the command line. No document
content changes: what changes is that an ambiguous invocation is refused instead
of silently doing something else.

### Changed

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
  was never created and the *input* file was overwritten instead. `-o` with
  `--dry-run` stays legal, and a run with none of the three landing flags still
  exits `2`. `insert_formula.py` follows the same rule for `--dump` with
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

### Notes

- `docx_mathify.py --inplace --dry-run` is still tolerated rather than refused:
  `--dry-run` wins and no backup is written. It is the one remaining silent
  combination, recorded so nobody relies on the reverse.

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

[1.0.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.0.0