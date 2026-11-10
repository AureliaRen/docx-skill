# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

- 65 tests across four new or extended files:
  - `tests/test_docx_edit_subcommands.py` — all eight `docx_edit` subcommands,
    including cross-run `replace` and a `normalize` case that asserts the merged
    runs keep every character.
  - `tests/test_docx_pack_safety.py` — drive-prefix, POSIX-absolute and relative
    traversal, each asserting both the refusal and that nothing was written.
  - `tests/test_docx_template_output.py` — `-o` honoured, source untouched,
    `.bak` on in-place writes.
  - `tests/test_insert_formula.py` — the `--display`, `--dump` and
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