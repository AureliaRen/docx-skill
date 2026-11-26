# Changelog

All notable changes to this project are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.3.2] - 2026-11-26

A silent-delivery round. LaTeX commands that OMML has no structure for - `\eqno`,
`\nonumber`, `\textwidth`, `\hline`, `\newpage` and 43 more - used to reach the document as
**literal text inside the equation**: `latex2mathml` emits an unsupported control
sequence as plain text instead of failing, `MML2OMML.XSL` copies it into an `<m:t>`
run, and Word draws the backslash and the command name where the equation should
be. Nothing in the toolkit objected. Measured on a four-paragraph document
carrying three such commands, with the new guards patched out: `docx_mathify`
printed `失败 0` and exited `0`, `docx_validate` returned `ok: true` with no
residue key, and the file held three backslash-only `m:t` runs. Every check was
green while the deliverable was visibly wrong - a defect in what the toolkit
*reported*, not a feature it lacked.

### Added

- **A pre-emptive control-sequence guard, `assert_no_unsupported_commands()`.**
  Runs at the head of `latex_to_omml_lxml`, before `keep_spacing`, and refuses the
  48 LaTeX-only names OMML cannot express (`LATEX_ONLY_COMMANDS`), each with the
  replacement to write instead - numbering to `\tag` or the numbering segment,
  cross-references to `docx_crossref`, page and column dimensions to the `page`
  block, spacing and breaks to paragraph formatting, table rules to border
  settings, macros to expanded text. Refusing at the source is the only place a
  message can explain **why**: once the command has become text, nothing
  downstream knows what it was meant to do.
- **Five position-dependent names on the same guard.** `\limits`, `\nolimits`,
  `\middle`, `\def` and `\newcommand` are legal in one position and leak in another, so
  `CONTEXT_RULES` decides by looking around the command instead of at its name:
  `bigop_prefix` (immediately after one of the 25 `BIG_OPERATORS`), `left_right`
  (inside an unclosed pair of `\left` and `\right`), `macro_use` (defined and then
  referenced). A name-based blacklist cannot do this: measured on one build, a
  subscripted `\sum` with `\nolimits` converts cleanly while `x\nolimits` leaks, and a
  used `\newcommand` yields `['1']` while an unused one yields an **empty** formula.
- **A post-conversion residue guard, `assert_no_latex_residue()`.** Runs after
  `split_runs_for_upright` so it scans the final tree, and reports any command
  that survived as literal text - including ones nobody has catalogued, which is
  the whole reason it exists next to the name list. The criterion is that the
  whitespace-separated token *is* a command; measured against three candidates it
  is the only one that catches text-mode residue (5 of 5) without flagging a
  legitimate path-bearing formula (0 of 3). Whole-run anchoring misses every
  text-mode leak because the command shares its run with the surrounding words;
  substring matching flags `\text{C:\Users\me}`. It works because `str.split()` splits
  on NBSP, which is how `latex2mathml` spaces text mode.
- **`docx_validate` reports `omml_latex_residue`.** `<m:t>` runs whose text is a
  command name are counted into one `warning` issue carrying `count` and the first
  three distinct `samples`. `<w:t>` is never scanned - a backslash in body text is
  legitimate content. Severity is `warning` on purpose: this check inspects a
  document that may already be on someone's disk, so `ok` and the exit code stay
  put and a clean document does not grow the key. Both sides share the shape test
  but not the consequence: the converter may refuse, the health check may not.

### Changed

- **`docx_mathify` failures carry a kind.** `latex_guard` means the source text
  holds a command OMML cannot express (fix the document); `convert` means the
  converter raised (check the environment - a missing `MML2OMML.XSL`, a syntax
  error, illegal structure). The two used to arrive in one undifferentiated list,
  which sent readers to reinstall Office over a `\eqno` and to edit a formula over
  a missing stylesheet.
- **The summary line splits the failure count.** `失败 N（拒收 M，转换 K）` - the total
  is unchanged, so `N` still matches the exit code, and the parenthetical names how
  many are which. It appears **only when there is at least one failure**: a clean
  run still prints `失败 0` and nothing else, byte-identical to the previous
  release, so anything parsing that line keeps working.
- **`FAIL[guard]` / `FAIL[convert]` prefixes on stdout**, up to ten lines of each
  kind. The prefix *is* the classification - no extra section, so the count in the
  summary and the count of printed lines cannot drift apart.
- **`--report` writes two sections**, one per kind, each naming the repair it
  wants. Mixed together, the reader takes a missing-XSL error to the formula and a
  `\eqno` to the installer.

**No new exit code.** D-024 stands: a document with at least one failed formula is
still `3`, both kinds included, and the output file is still written. The kind is
carried by the prefix on stdout and by the failure record, never by the process
status.

## [1.3.1] - 2026-11-26

A delimiter round. A bracket group whose first item is a unary sign —
`\left(-\frac{y}{x}\right)`, and its bare-bracket twin `(-\frac{y}{x})` —
converted into an `m:d` with an empty first slot, which Word draws as an empty box
inside the brackets while every check in the toolkit stayed green. The cause sits
in Microsoft's `MML2OMML.XSL`, so the repair went in upstream of it: a MathML pass
that regroups the sign with its operand before the transform. Both structural
readers now see the condition, and the manual gained the section that was missing.

### Added

- **`assert_legal()` refuses a delimiter with an empty slot.** A third structural
  check next to the `m:oMath`-nesting and `m:rad` cases: an `m:d` owning an empty
  `m:e` now raises `RuntimeError` instead of shipping a blank square into the
  document. The message splits the diagnosis in two: the sign-at-head shape, which
  `wrap_leading_unary_in_delim` already repairs on the normal path (so hitting it
  means the formula bypassed `math_omml`), and a genuinely empty group
  (`()` with nothing between the delimiters), which needs content — never a
  different bracket spelling. Every conversion path runs through `assert_legal()`,
  so a formula cannot reach a document unnoticed.
- **`docx_validate` reports `omml_empty_slot`.** Delimiter groups in
  `word/document.xml` that own an empty `m:e` are counted and reported as one
  issue carrying `count`. Severity is `warning` on purpose: the report feeds
  pass/fail decisions downstream, and a document Word opens happily must not be
  rejected by a health check, so `ok` and the exit code are unchanged and a
  document with no `m:d` at all never grows the key. This is what catches work
  produced before the repair existed, or by another tool entirely.
- **`stats["delim_wraps"]`** — the number of groups the MathML pass rewrapped,
  accumulated through `latex_to_omml_lxml` → `latex_to_omml_safe` next to
  `upright_split` and `cases_aligned`. `insert_formula --dump` still prints no
  stats at all; its stdout stays parseable markup.

### Changed

- **`SKILL.md`'s output-convention list now names `pdf_export.ps1`, not
  `pdf_export`.** The list sits beside a dozen pure-Python CLIs, so a bare name
  read as `python scripts/pdf_export.py` — a file that does not exist, and the
  run dies on `can't open file`. It now carries the extension and says it is
  PowerShell (`pwsh`), matching the four other places in both manuals that
  already spell out `pwsh scripts/pdf_export.ps1`. `SKILL.zh-CN.md` carries the
  same list and the same fix. The file counts are untouched: `scripts/` is still
  28 `.py` + 1 `.ps1` = 29 files.
- **The README test readings were refreshed, and the test-module counts were
  corrected.** The developer row is this round's run (2607 passed / 239 skipped,
  2846 collected, 2026-11-26); the two runner rows keep their measured PR #6
  figures and now say they predate this round's 68 new tests, since no runner has
  run them yet. `tests/` is **84 `.py`** (83 `test_*.py` plus the `_omml.py`
  helper) after this round added `test_math_delim_slots.py` — `README.md`,
  `README.zh-CN.md`, both `SKILL` files and `THIRD-PARTY-NOTICES.md` all said
  83/82 and now agree with the tree.
- **`/docs/` is ignored by rule, not by habit.** Round reports carry machine-local
  paths, the account name and real coursework filenames; every release so far kept
  them out by typing an explicit `git add` list. `.gitignore` now anchors
  `/docs/` next to `/work/`, so `git add -A` cannot pick them up by accident.

### Fixed

- **A unary sign at the head of a delimiter group produced an empty `m:e` — an
  empty box inside the brackets in Word, and in the PDF exported through
  `pdf_export.ps1`.** `\left(-\frac{y}{x}\right)` converts to
  `<m:d><m:dPr><m:sepChr m:val="−"/></m:dPr><m:e/><m:e>…</m:e></m:d>`: the
  stylesheet reads the leading sign as the slot separator, so the slot in front of
  it stays empty. The equation is otherwise correct — which is why it survived
  review: `docx_mathify` reported `失败 0` and exit `0`, and
  `docx_validate` returned `{"ok": true, "issues": []}`.
  **The cause is in `MML2OMML.XSL`, not in this skill**: feeding
  `latex2mathml`'s MathML straight into `math_omml.xslt()`, with no skill code in
  the path, reproduces the same element. **The fix is therefore upstream of the
  stylesheet**: `wrap_leading_unary_in_delim()` (`math_omml.py`) walks the MathML
  tree before the transform and, when a group opens with a unary
  `−` / `-` / `+` / `±` / `∓` and holds at least one more item, wraps that sign
  plus the next sibling into a single `mrow`. The transform then sees one item
  where it saw two, emits no separator, and the group becomes one non-empty
  `m:e`. Nothing downstream rebuilds runs or re-copies a `w:rPr`.
  **The trigger is wider than a round bracket, and wider than `\left`:**
  `\left(-x\right)`, `\left(-x y\right)`, `\left(-1\right)`,
  `\left(-\frac{y}{x}\right)`, `\left(+\frac{a}{b}\right)`,
  `\left(\mp x\right)`, `\left[-x\right]`, the `\left\| … \right\|`
  absolute-value form and `\left\{-x\right\}` all measured one empty slot; the
  bare-bracket forms `(-x)`, `(-1)`, `(-\frac{y}{x})`, `[-x]`, `\{-x\}` and
  `\|-x\|` measured the same, because `fence="true"` is never read by the
  stylesheet. There is no bracket style to switch to, so nothing is rewritten by
  hand: both spellings broke identically, and the same pass clears both. It is not
  every leading sign either:
  `\left(-x+y\right)`, `(-\frac{a}{b}+1)` and `(-x,\ y)` were already free of
  empty slots, because a `+` or `,` behind the sign becomes the separator instead.
  **What did not change, and what only changed shape:** `\left(a+b\right)` keeps
  `sepChr='+'` and two slots and `\left(a,\ b\right)` keeps `','` — the pass never
  touches a group that opens with an ordinary symbol. The three clean rows above do
  get rewrapped, structure only: each held one slot with an empty `sepChr` and now
  holds two separated by `'+'` / `'+'` / `','`. Word draws `(−x+y)`, `(−a/b+1)` and
  `(−x, y)` identically either way; an empty separator is not a rendering
  feature. A genuinely empty `\left(\right)` still measures one empty slot, since
  there is no second item to pair a sign with; it is now rejected by
  `assert_legal()` rather than shipped.
  Both manuals carry a new section on the symptom, the trigger matrix, the
  reproduction, the self-check and the boundaries; `references/formula-guide.md`
  carries the long form.

## [1.3.0] - 2026-11-25

A formula-layout round. Word draws letters in a math zone italic because a lone
letter is a variable, so function names and differential operators arrived
slanted: `\max\{x,y\}` came through as one run of literal `max{x,y}`, and
`\frac{dy}{dx}` set the `d` italic along with the `x`. `cases` rendered
centre-aligned while every matrix looked fine, and nothing in the OMML told the
two apart. The fixes are two passes inside the converter, plus the documentation
that says what each renderer can and cannot verify.

### Added

- **Upright completion for function names and differentials.**
  `split_runs_for_upright()` (`math_omml.py`) classifies each `m:r`'s characters
  and rebuilds the run as segments, giving digit runs, operators and the 34
  whitelisted names in `UPRIGHT_FUNCS` their own `m:rPr/m:sty val="p"`. It runs
  inside `latex_to_omml_lxml()` after `apply_math_font()` — the split has to
  `deepcopy` the already-injected `w:rPr`, or the new segments fall back to
  Cambria Math mid-formula — and `assert_legal()` re-runs afterwards, because
  rebuilding runs is itself a structural edit. Deliberately not a switch: this is
  layout correctness, not a preference.
- **`\begin{cases}` columns are left-aligned.** `align_cases_left()` reads the
  MathML tree the converter still holds (a `cases` `mtd` carries
  `columnalign="left"`, a matrix's carries none) and rewrites only the matching
  `m:m` to `mcJc=left`. Positional, so the counts must agree first: only tables
  with two or more columns produce an `m:m`, and when the counts disagree it
  changes nothing and records `stats["cases_align_skipped"]`. A blanket
  `mcJc=left` pass would left-align every matrix in the document.
- **Statistics for both passes.** `upright_split` / `upright_runs` /
  `cases_aligned` / `cases_align_skipped` accumulate into the caller's `stats`
  dict, from `latex_to_omml_lxml` through `latex_to_omml_safe`.

### Changed

- **The manual no longer asks for `\mathrm{max}` as the only correct spelling.**
  Now that the converter uprights `max` on its own, `\max\{x,y\}` and
  `\mathrm{max}\{x,y\}` produce segment-for-segment identical OMML and render to
  the same pixels in Word (measured: 0 differing pixels; before the pass the two
  forms differed in several hundred to several thousand pixels depending on the
  rasterisation DPI — 371 at 72 dpi, 1195 at 150 dpi on a 1275x1650 page). The
  recommendation stands — `\mathrm{}` states the intent to every other renderer —
  but it is no longer a correctness requirement, and the documentation says so
  instead of calling the plain form wrong.
- **Documented: LaTeX literals in Python are raw strings, and why it matters.**
  32 common commands lose their backslash to a control-character escape inside an
  ordinary literal (`\frac`, `\forall` via `\f`; `\times`, `\text`, `\theta`,
  `\tan`, `\top` via `\t`; `\nabla`, `\neq`, `\nu` via `\n`; `\beta`, `\bar`,
  `\begin`, `\bmod` via `\b`; `\rho`, `\right`, `\ref` via `\r`; `\alpha`,
  `\approx`, `\angle`, `\array` via `\a`; `\vee`, `\vec`, `\varphi` via `\v`),
  and four more fail at parse time. Five of the seven escape classes then convert
  **silently** into a wrong formula (`rac12`, `2imes3`, `hoV`) rather than
  raising — that asymmetry is what makes it a hard rule. JSON doubles the
  backslash, so a `docx_create` spec writes `"latex": "\\frac{a}{b} = c"`.
- **Documented: body font and equation font are separate layers.** The equation
  layer lives inside each `m:r` and is written by `apply_math_font()` from
  `MATH_FONT`, across all four of `ascii`/`hAnsi`/`cs`/`eastAsia`. Switching the
  body's CJK font must leave the equation's `eastAsia` alone, or `\text{中文}`
  stops matching the prose around it while every Latin glyph still lines up. Change
  the equation font through `MATH_FONT`; `fix_fonts.py` already honours the split.
- **Documented: what each renderer can verify.** LibreOffice ignores `m:sty`
  entirely — an italic and an upright `max` rasterise to the same pixels through
  `docx_preview.py` (measured XOR 0 at every DPI tried; through `pdf_export.ps1`
  the same pair differs by hundreds to thousands of pixels depending on DPI) — so
  preview cannot confirm or refute the upright pass. Glyph shape, integral size and
  delimiter height are decided on the Word-exported PDF only; preview is for
  "did the equation get created, in the right paragraph, still inline". The
  per-feature pixel measurements and the three-step division of labour are in
  `references/formula-guide.md`.
- **Documented: `$...$` is single-dollar only, and `$$...$$` fails silently.**
  The doubled form is not a display marker and is not recognised at all: the
  detector lifts the LaTeX out of the run and converts it correctly, every `$` the
  author typed stays in the prose, and the run still reports `失败 0` with exit
  `0`. Measured on `docx_mathify`: `由 $$\sigma=\frac{F}{A}$$ 可知` comes back as
  `由 $$$$ 可知` with `oMath=1`. Checking it therefore means reading `w:t` back and
  grepping for a residual `$`; a display equation is a `docx_create` `latex` block
  with `"display": true` or `insert_formula.py --display`. Also recorded that the
  shipped CLIs carry no marker layer — `$x=1$` in a paragraph stays literal text —
  so the `$...$` loop in `formula-guide.md` is a recipe for your own caller, and a
  currency `$` in prose needs no escaping.
- **README test readings refreshed** against this round's runs: 2539 passed / 239
  skipped / 0 failed on a developer machine with Word installed, and 2233 passed /
  545 skipped (ubuntu) vs 2236 / 542 (windows) on the four CI legs, all 2778
  collected, dated 2026-11-25. The `ci.yml` floor moves 2231 -> **2233**, the lower
  of the two runner legs — floors only move up when a real runner run says so.
- **Documented: `cases` and matrices are the same object in OMML**, and a scanned
  textbook's printed page number is not the PDF page index — locate a formula by
  its equation number or section heading, never by offsetting the printed number.
- **Test-suite counts corrected.** `tests/` is 83 `.py` (82 `test_*.py` + the
  `_omml.py` helper), not 81/80; `pytest --collect-only -q` reports 2778 collected
  on 2026-11-25. Both `SKILL.md` and `SKILL.zh-CN.md` had the old figures.

## [1.2.1] - 2026-11-24

A documentation-and-entry-point round driven by a real delivery that hit five
walls. Two of them were code: there was no supported way to write a document
whose equations start as LaTeX, and a hand-written caller could skip the
phasor-dot repair while every documented example looked fine. The other three
were the manual telling the truth late: no section on importing the scripts from
your own code, no warning that preview output is LibreOffice rather than Word,
and a stdio note that described the intent rather than the measured behaviour.

### Added

- **A `latex` block in a `docx_create` spec.** `{"type": "latex", "latex":
  "\\frac{a}{b}=c", "display": true}` writes one native Office Math equation:
  required non-empty `latex`, optional boolean `display` (default `false`), the
  usual paragraph format keys, and the block converts through
  `latex_to_omml_safe()`. An inline equation leaves a U+200B run beside it so
  Word does not promote a paragraph-wide inline equation to a display one.
  `display` that is not a real boolean is a spec error (exit `2`, nothing
  written). A formula that fails to convert does not abort the document: the
  LaTeX source stays in the paragraph, `formula_failures` lists it on stdout,
  one `warning:` line goes to stderr, and the exit code is `3` — the JSON says
  `"ok": false` while `output` still names the file that exists, so branch on
  `output`.
- **`math_omml.latex_to_omml_safe()`** — conversion + `fix_phasor_dots` + a
  second `assert_legal` in one call, with optional `stats` counters. Hand-written
  callers no longer have to remember the phasor-dot pass. `latex_to_omml_lxml`
  stays as the pure converter and now opens its docstring with a warning that it
  repairs nothing.
- **Documentation: importing the scripts from your own code.** `scripts/` is
  flat, so a sibling import resolves only when that directory is on `sys.path`;
  a script of your own hit `ModuleNotFoundError` with nothing in the manual to
  explain it. There is now a section with both copyable forms (`sys.path.insert`
  and `import _bootstrap`), a runnable pair, and a new walkthrough for writing a
  document that contains formulas from scratch — spec route and hand-written
  route, with the three behaviours that surprise people (the zero-width run,
  one failed formula not aborting the run, the `latex2mathml` + `MML2OMML.XSL`
  requirement and the fact that a missing closing brace still converts).

### Changed

- **Windows stdio hardening is verified, not assumed.** `encoding._fix_stdio()`
  now reconfigures to UTF-8 with `errors="backslashreplace"`, re-reads the
  stream's codec-normalised encoding and `errors` (a stream still on `replace` /
  `ignore` counts as unfinished), and rebinds through `io.TextIOWrapper` only
  when that check fails; non-Windows is untouched. `STDIO_MODE` / `STDIO_MODES`
  expose which branch ran, and `tools/word_accept.py` carries the same policy in
  its own standalone copy.
- **The manual no longer claims preview output equals Word's.** `docx_preview.py`
  renders through headless LibreOffice, so line breaks, page fill and sometimes
  page count differ from a Word COM export; the delivered layout is judged from
  `pdf_export.ps1`'s PDF, and preview is for a fast look plus the structural
  lint.

## [1.2.0] - 2026-11-23

A patch round that is almost entirely about *silent* damage: four ways an edit
could finish with `exit 0`, an empty stderr and `{"ok": true}`, and still leave
the document wrong. `set-cell` was the worst of them - filling a template cell
flattened its runs and dropped sub/superscript and alignment without a word.
Reading and writing now agree on what a run is, `--replace-cell` keeps the old
clearing behaviour for callers who want it, and the validator finally looks
inside `<w:drawing>`.

The same round closes a fifth silent path on the formula side: a vulgar
fraction character (`½`, `⅓`, `⅕`…) could survive a conversion as a raw glyph.
Fourteen of the nineteen fraction code points were not even recognised as math,
and the five that were still had to justify themselves with an `=` somewhere in
the same run. `x=½` converted, `½+1` did not - same character, two outcomes.

Chasing that one turned up a sixth, older and nastier: a spacing command was two
tokens, so `1/2 m/s` came out with a denominator reading `2},`. Word opens it,
it looks like a fraction from across the desk, and nothing anywhere reports it.

A second round widened the same contract to the places the first could not
reach. `docx_preview.py` decoded soffice's UTF-8 output with the console
locale, so on a cp936 machine a document with CJK text raised
`UnicodeDecodeError` inside a pipe-reader thread that Python silently reaps -
`exit 0`, empty stderr, and a crash nobody logged. The reader could not see
text boxes at all: `docx_read.py` walked only the body's direct `p`/`tbl`
children, so 43 text boxes in a real annotated assignment read as zero, in
every mode. And the validator judged run-property values by the XML schema
instead of by Word: `darkMagenta` opens, `DARK_MAGENTA` does not, and both
passed.

### Changed

- **`docx_edit.py set-cell` keeps cell formatting by default** (was: clear and
  rewrite). Filling a template cell used to go through `_Cell.text`, whose
  setter is documented as "Any existing content or revisions are replaced": it
  clears the cell, adds one paragraph and one run, and drops paragraph
  alignment, run font, run size, bold/italic and `vertAlign`. Measured on a
  1x1 table with 2 paragraphs and 3 runs (second run `subscript`, first
  paragraph `CENTER`, Times New Roman 14 pt), writing `"d0=5"` left **1
  paragraph, 1 run, `CENTER` -> `None`, 14 pt / TNR -> `None`** - and reported
  `{"ok": true}`, exit `0`, empty stderr. The new default is `mode: "preserve"`:

  | Situation | Behaviour | `mode` |
  |---|---|---|
  | One writable run | Replace its text, leave `rPr` alone | `preserve` |
  | Several runs | Write into the first writable run, drop the rest, strip a stray `vertAlign` from that run, **warn** | `preserve` + `warnings` |
  | `--replace-cell` | Exactly the old clearing behaviour | `replace-cell` |

  A run is not writable when it is a field (`w:fldChar` / `w:instrText`), a
  drawing, or a picture - those are kept rather than clobbered, because
  replacing a field run breaks the result it produced. Whatever could not be
  kept is reported on stderr as a `warning:` line and in the JSON `warnings`
  array: `exit 0` no longer implies the cell came out identical.
- **`references/` is 7 files, not 6** - `editing-traps.md` joins it: the six
  silent-failure traps (`cell.text =` flattening, stray `vertAlign`,
  `inline_shapes` index drift, empty `drawing`, `PermissionError` from an open
  WPS/Word file, cp936 mojibake), each with the measured before/after table and
  the correct code shape (`copy.deepcopy` on the `w:r` element, matching images
  by `rId` rather than position).
- **All nineteen vulgar fraction characters convert to `\frac{n}{d}`** (was:
  five of nineteen, and only when the run also contained a relation operator).
  `½ ¼ ¾ ⅓ ⅔` were in the symbol table; `⅐ ⅑ ⅒ ⅕ ⅖ ⅗ ⅘ ⅙ ⅚ ⅛ ⅜ ⅝ ⅞ ↉` were not
  even math characters, so they stayed in body text as glyphs forever. The
  glyph is a hazard rather than an error: U+00BD renders as a tiny raised `1`,
  a slash and a lowered `2`, which looks *nearly* right in one font and wrong
  in every other - nobody notices until print. Measured before: `½` -> `½`
  (unchanged), `½+1` -> `½+1`, `x=½` -> `x=\frac{1}{2}`; the third converted
  and the first two did not, from the same character.

  A fraction character is now its own justification for opening a math span, on
  the same footing as `√`/`∑`/`∫` and as an explicit `_{...}` group: it is a
  math glyph the author typed deliberately, not a paragraph we guessed at. It
  also lifted the `len >= 2` floor for that one case, so a lone `½` between
  Chinese characters converts while all framing stays text.
- **A fraction character is a value anchor for unit strings**, so `½ m/s` reads
  as a measurement (`\frac{1}{2}\,\text{m/s}`) exactly like `0.5 m/s` does
  (`0.5\,\text{m/s}`) - same quantity, same layout. Without the anchor it went
  through the fraction rewriter instead and came out as
  `\frac{1}{2} \frac{m}{s}`, i.e. "half divided by metres per second". Two
  adjacent lines of one hand-out disagreeing is harder to spot than either
  being wrong. The unit grammar itself is untouched: `kg/m³`, `J/(mol·K)`,
  `N·m` and `8.31 J/(mol·K)` are byte-identical before and after.
- **A spacing command is one token, so a fraction no longer swallows the unit
  after it** (was: `1/2 m/s` produced malformed LaTeX). `\,` is two characters
  and contains no letter, so `_CMD_RE` (`\[A-Za-z]+`) could not match it, and
  `_split_atoms` fell through to its "backslash is a single atom" branch — the
  comma became an *operator* and the bare backslash became an *adjacent factor*.
  With a fraction and a unit string in the same span, which is what a hand-out
  writes constantly:

  ```
  1/2 m/s  → 1/2\,\text{m/s} → \frac{1}{2\},\text{m/s}   ← denominator is 2\}
  ```

  The escaped brace closed the group early: OMML came out with numerator `1`,
  denominator `2},`, then `m/s`. Word draws a stacked fraction whose denominator
  reads `2},`. It opens fine and looks like a fraction from across the desk, so
  nobody reported it. The five two-character spacing commands (`\,` `\;` `\:`
  `\!` `\ `) are now collected as a single `space` token before any operand
  scanning runs. Measured after: `1/2 m/s` → `\frac{1}{2}\,\text{m/s}` (denominator
  `2`), `3/4 m/s`, `1/2 rad/s`, `1/2 J/(mol·K)`, `1/2 kg/m³` all well-formed and
  brace-balanced. `1/2×1/2` still refuses to nest and `v=m/s` is still a
  fraction, not a unit string.

### Added

- **`docx_read.py --runs`** - per-paragraph run inventory: for each run its
  `text`, `vert_align`, `bold`, `italic`, `size_pt`, `font`, `is_field` and
  `is_drawing`, plus a paragraph-level `mixed_format` flag and a `path` such as
  `table[0].r0.c0`. `mixed_format: true` is the signal that a flat string
  cannot express the cell. Read this **before** writing into a template cell.
  `--blocks` is unchanged - its fields are consumed by the Markdown round-trip
  path, so the run detail went into a new switch rather than a new field.
- **`docx_validate.py` rejects an empty `<w:drawing>`** - code `empty-drawing`,
  severity `error`, exit `1`. A drawing must contain exactly one `wp:inline` or
  `wp:anchor`; Word does not repair one that contains neither, it refuses to
  open the file. Measured before the check existed: `{"ok": true, "issues":
  []}`, exit `0` - and `Documents.Open()` raised `com_error (-2147352567,
  '发生意外。', ..., 'Word 在试图打开文件时遇到错误。...', 'wdmain11.chm', 24601,
  ...)`. The five earlier checks (package structure, rels, referenced ids,
  media magic, style existence) all pass an empty drawing, because it
  references nothing. Verified against five real documents (`in.docx`,
  `out.docx`, `out_tnr.docx`, `final_test.docx`, `fixture.docx`) - all five
  still `{"ok": true, "issues": []}`.
- **`docx_validate.py` rejects a block element nested directly inside a run** -
  code `invalid-nesting`, severity `error`, exit `1`. A `<w:p>` or `<w:tbl>`
  inside `<w:r>` is not valid run content: Word either silently repairs it or
  refuses the file, while python-docx opens it happily, so no other check here
  would notice. Only those two tags are flagged - narrowing to the two shapes
  is deliberate, because a wider net starts reporting legal `mc:AlternateContent`
  fallbacks.
- **Malformed XML in `docx_validate.py` is a JSON verdict, not a traceback** -
  a document whose XML does not parse now produces `ok: false` with the parse
  error as an issue (`bad-document-xml` for the body, `bad-styles-xml` for the
  style part), instead of dumping a stack trace and losing the exit code
  contract.
- **`docx_validate.py` no longer reports phantom unresolved references** (a
  pre-existing bug, and the worst kind: it depended on `PYTHONHASHSEED`).
  Every `.rels` file used to be keyed by its *parent directory*, so
  `word/_rels/document.xml.rels`, `header1.xml.rels`, `footer1.xml.rels`,
  `comments.xml.rels` and friends all collapsed into one `"word"` slot and the
  last one read won. Since the file list is a `set`, which one won varied per
  process. Measured on `source-snapshot/test_unpack/full_test_input.docx`
  against the previous commit: seeds `0,3,4,5,7` reported two
  `unresolved-reference` errors while `1,2,6` reported `{"ok": true, "issues":
  []}` - same file, same code, different answer. The key is now the owning part
  (`word/document.xml`), resolved by file stem rather than by directory, and the
  `.rels` files are iterated in sorted order so the result cannot depend on
  hash order. All eight seeds now agree.
- **`docx_toc.py --levels 0` is refused** (was: exit `0` with an empty TOC of
  zero entries). An outline of nothing is a mistake, not an answer; the refusal
  names the accepted forms (`1,2` / `1-3`). Non-numeric `--indent-chars` now
  fails as a single `error:` line instead of a bare traceback.
- **`docx_read.py` reads text boxes** - `w:txbxContent` is enumerated wherever
  it sits (drawn shape, VML pict, or `mc:AlternateContent`), with `mc:Choice`
  preferred over its `mc:Fallback` twin so each box is reported once, not
  twice. `--blocks` gains `textbox` blocks with a `flavor` field, `--text`
  gains a `textboxes` key, `--runs` tags textbox paragraphs, `--md` renders
  them as quoted blocks. Measured on a real 43-box document: every mode
  reports 43 (the naive count is 86 - Choice and Fallback each carry a copy);
  an independent lxml walk agrees, 43/43/43, zero empty branches.
- **`docx_edit.py char-format`** - set `--highlight`, `--color`, `--bold`,
  `--italic` on every run matching `--find`. The 17 legal highlight values
  accept the names Word itself writes (`darkMagenta`) and the enum spellings
  (`DARK_MAGENTA`, `dark_magenta`); anything else is a single `error:` line
  and exit `2` - because Word refuses the file, not the style: measured with
  Word COM, `darkMagenta` opens while `DARK_MAGENTA` / `purple` / `FF0000`
  all raise the corrupt-file dialog (com_error 25272). Matching runs are
  split and deep-copied, so text before and after the match keeps its
  formatting. `docx_read.py --runs` now reports `highlight` and `color` per
  run.
- **`docx_validate.py` audits run-property values** - two new codes:
  `invalid-run-property` (error - a value Word refuses to open the file for:
  `w:highlight`, `w:vertAlign`, `w:effect`) and `nonstandard-run-property`
  (warning - a value Word tolerates but reinterprets: `w:u`, `w:color`).
  Severity comes from measured Word behaviour, not from the schema:
  `vertAlign="normal"` opens fine, while `effect="ants"` is schema-legal and
  Word still refuses it. Nine parts are scanned (document, styles,
  stylesWithEffects, numbering, settings, comments, footnotes, endnotes,
  headers/footers). Zero false positives across a 407-document healthy
  corpus.
- **Element order has one source** - the run- and paragraph-property
  sequences live in `docx_utils._ORDER_MAP`, consumed through a
  `successors()` helper in `docx_crossref`; the three hard-coded copies in
  `docx_create.py` are gone (two of them had quietly drifted into paragraph
  order). 40 parts re-saved byte-identical before and after the switch.

### Fixed

- **`docx_preview.py` reads soffice output as UTF-8 explicitly** -
  `encoding="utf-8", errors="replace"` on the one bare `text=True`
  subprocess call in production code. Measured: soffice emits UTF-8; under a
  cp936 locale (`PYTHONUTF8=0`) the old call died on byte `0xaf` inside the
  reaped reader thread - `exit 0`, nothing logged. Eight test-side calls were
  measured safe under the same locale and left alone.
- **README / SKILL / THIRD-PARTY-NOTICES re-aligned with the tree** - counts
  had drifted three phases behind (scripts 27 -> 29, tests 47 -> 77,
  references 6 -> 7), and the Hermes lineage list named `docx_common` and
  `docx_template` as header-marked while the markers actually sit on
  `docx_preview` and `docx_footer` (with `docx_template`'s evidence in its
  docstring and `docx_common`'s in its family header). The notices table now
  lists all fourteen derived files plus the two test modules carrying the
  same marker.

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

- **An `image` block with a bad `path` is an input error, not a traceback.** The
  path went straight into `add_picture`, so a typo reached the user as a
  `FileNotFoundError` raised from inside the build: a traceback, no JSON error
  object, and no block index anywhere in the message. It is now checked with the
  other input errors — exit `2`, one `error:` line naming `blocks[N]`, nothing
  written — and the message says the path resolves against the current working
  directory rather than the spec file's directory, which is where the first
  report of this lost ten minutes.
- **`docx_preview.py` names the first blocker, not the second.** PyMuPDF was
  checked before LibreOffice even though conversion runs first, so a machine with
  neither — every GitHub runner — was told to install PyMuPDF and then hit the
  same wall one step later. The checks now run in execution order.
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

[1.3.2]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.3.2
[1.3.1]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.3.1
[1.3.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.3.0
[1.2.1]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.2.1
[1.2.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.2.0
[1.1.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.1.0
[1.0.0]: https://github.com/AureliaRen/docx-skill/releases/tag/v1.0.0