# The formula chain

Four modules cooperate. Understanding the split tells you where to look when
something goes wrong.

```
docx (zip)
   │
   ├─ linear_to_latex.py    decide WHAT is a formula, render it as LaTeX
   │      find_math_spans()      which atom ranges count
   │      linear_to_latex()      linear notation -> LaTeX text
   │      repair_merged_scripts() fix "I^{2R}" -> "I^{2}R"
   │
   ├─ math_omml.py          LaTeX -> MathML -> OMML
   │      latex_to_omml_lxml()   returns an lxml element holding <m:oMath>
   │      assert_legal()         structural self-check
   │      apply_math_font()      inject w:rPr/w:rFonts
   │
   ├─ insert_formula.py     one equation, at an anchor, into an unpacked doc
   │
   └─ docx_mathify.py       batch: scan, convert, place, report
          _om_or_scripts()       the tier split (see below)
          no_layout_reason()     tier 3: does this need an equation object at all
          is_pure_script()       tier 2: scripts but no structure
```

## Three output tiers

Recognising a span as a formula is one decision; deciding what to emit for it is
another. `_om_or_scripts()` in `docx_mathify.py` makes the second decision in a
fixed order, and each tier has its own test:

| Tier | Test | Output | Example |
|---|---|---|---|
| 1. Structural | `structural_reason()` finds a command in `_STRUCTURAL_CMDS` — 43 entries covering fractions, radicals, integrals, big operators, limits, accents, delimiters, environments, and equation tags | OMML via `latex_to_omml_lxml()` | `x^2/16+y^2/9=1`, `v=s/t` |
| 2. Plain script | no structural command, and the span contains at least one `^` or `_` | ordinary `<w:r>` runs carrying `w:rPr/w:vertAlign` | `a=9.8m/s2`, `10^-7` |
| 3. No layout needed | no structural command, no `^`/`_`, exactly one `=`, a single ASCII letter (optionally with a trailing `'`) on the left, a signed decimal on the right, and an optional unit tail | ordinary text runs, original `rPr` preserved | `s=10km`, `t=2h`, `m=5kg`, `v=3m/s` |

The order matters and is not interchangeable. Tier 3 requires the *absence* of
`^`/`_`; tier 2 requires their *presence*. They are mutually exclusive by
construction, and `no_layout_reason()` is a pure function so the switch is
applied only at the call site, never inside the test.

### Why tier 2 exists at all

Word stores a real superscript as `<w:rPr><w:vertAlign w:val="superscript"/>`,
which is what `Ctrl+Shift+=` toggles and what restores the character to full size
when toggled back. An OMML `m:sSup` cannot do that: its raised glyph is a
formula-engine layout, `Font.Superscript` reports `0`, and pressing the shortcut
on it stacks a second superscript on already-shrunken type. Tier 2 emits the run
form so the shortcut behaves the way a Word user expects.

### Why tier 3 costs the italic, and why that is settled

Tier 2 and tier 3 both lose the math italic that OMML would have supplied. For a
span that genuinely needs scripts, that is the price of getting real scripts. For
tier 3 the accounting was revisited and the cost accepted on 2026-11-07: where
nothing needs two-dimensional layout, a thing expressible as text should not
become an equation object.

The reasoning is about the object, not the glyph. An equation object in a body
paragraph means opening the formula editor to change one character, losing shape
when copied into other software, and a line box the formula engine stretches.
Against that, `s=10km` has no fraction, no radical, no script — nothing the
object is for. So the span is emitted as text and `s` comes out upright.

The correct remedy when the italic is wanted is to change the *input*, not to
push the span back up a tier:

- write the form with a script — `a=9.8m/s2` becomes tier 2 and keeps a real
  superscript run;
- or, building a document from a spec, mark the runs explicitly with
  `vert_align` (`superscript` / `subscript`; the `superscript` and `subscript`
  booleans are accepted as aliases, and an explicit `vert_align` wins over
  them) in `docx_create`.

`w:i` is not the answer: it is run-wide italic, so `18` and `g/mol` would be
slanted along with the quantity symbol.

### What tier 3 refuses, and why

The unit tail is a whitelist rather than "any letters". `_is_unit_tail()` splits
on `/`, `·`, `⋅` and requires every factor to be a member of
`linear_to_latex.UNITS`, or `%`. Two consequences worth stating outright:

- `y=2x` stays an equation. A permissive tail would read it as "2 metres" and
  flatten a function relation to upright text.
- `n=2t` stays an equation. `t` is deliberately absent from `UNITS` (it must not
  shadow time as a quantity symbol), so `v=s/t` keeps its real fraction, and
  `n=2t` is refused by the same rule.

Also refused, all conservatively — the span keeps its equation object:

| Input | Refused at | Reason |
|---|---|---|
| `F=ma`, `U=IR` | right-hand side | not a pure number |
| `f(x)=0`, `A=(-1,3)` | left-hand side | not a single letter |
| `n\geq 2` | `=` count | no equation sign |
| `m=\rho V=1.2\times 60=72kg` | `=` count | chained equality |
| `a=9.8\,\mathrm{m/s^{2}}` | `^`/`_` | tier 2 picks it up instead |
| `y=2x`, `n=2t` | unit whitelist | see above |
| `s=10m/`, `s=10/x` | unit whitelist | incomplete or non-unit factor |

Two current limits, recorded as behaviour rather than as guarantees: `T=25℃`
stays an equation because `℃` is not in `UNITS`, and a Greek quantity symbol such
as `η=80%` stays an equation because the left-hand side must be a single ASCII
letter. Both err toward keeping the equation object, so neither can mangle text.

### The two switches

Both are module-level constants in `docx_mathify.py`:

| Constant | Default | `False` |
|---|---|---|
| `MIXED_SUPERSCRIPT` | `True` | every span goes to OMML — the master switch |
| `NO_LAYOUT_NEEDED` | `True` | only tier 3 stops; tier 2 still applies |

`NO_LAYOUT_NEEDED` is contained by `MIXED_SUPERSCRIPT`. Turning the master switch
off stops both tiers, which is what makes "`MIXED_SUPERSCRIPT = False` restores
the old behaviour" a true statement rather than a nearly-true one. The report
recomputes the tier-3 reason only when both switches are on, so a report cannot
claim a downgrade that the document did not get.

## Why the conversion goes through Microsoft's stylesheet

The obvious approach — hand-building OMML XML — produces documents Word rejects
with "the file may be corrupted" (error 25272) or renders as a lone `∫`. The
failures are subtle schema violations: an `<m:oMath>` nested inside `<m:num>`,
an `<m:rad>` with no `<m:e>`, a `<m:nary>` that lost its `naryPr`. They pass a
naive eye and even some XSD checks.

Delegating to `MML2OMML.XSL` — the stylesheet Microsoft ships with Office — makes
the output correct by construction. Do not hand-build OMML.

`assert_legal()` remains as a tripwire for the two mistakes that hurt most:

- a `<m:oMath>` appearing inside any of
  `m:num` / `m:den` / `m:e` / `m:sub` / `m:sup` / `m:deg` / `m:fName` / `m:lim`
- a `<m:rad>` with no `<m:e>` child

That first list is the set `ARG_TAGS` in `math_omml.py` — `num`, `den`, `e`,
`sub`, `sup`, `deg`, `fName`, `lim`, eight argument containers. Each of them
holds runs and nested constructs: an `<m:e>` may contain an `m:f`, an `m:rad`,
an `m:nary`, all day long. What it may not contain is a whole equation, because
`m:oMath` is the outermost element of the construct, not an argument of it.
Word's reader parses the nesting without complaint and then declines to lay the
document out, which is why the failure surfaces to the user as "the file may be
corrupted" instead of as a parse error pointing at the offending byte. The
missing `<m:e>` fails the same way: a radical with nothing under it has no
layout, so the whole body is abandoned.

Both checks run inside `latex_to_omml_lxml()`, immediately after the stylesheet
returns and before the font injection, which is the only point where the tree is
still the stylesheet's output. `assert_legal()` quotes the LaTeX source in the
error, so the message names the input you wrote rather than a byte offset in a
part you never edited by hand.

## Locating MML2OMML.XSL

`math_omml.find_xslt()` checks, in order:

1. `$MML2OMML_XSL` — an explicit file path. Use this when Office lives somewhere
   unusual, or to point at a stylesheet copied out of a container.
2. `$ProgramFiles`, `$ProgramFiles(x86)`, `$LOCALAPPDATA`, then the literal
   `C:\Program Files` and `C:\Program Files (x86)`.

Under each root it globs:

```
Microsoft Office/**/MML2OMML.XSL
Microsoft Office/MML2OMML.XSL
*/*/Office*/MML2OMML.XSL
Kingsoft/**/MML2OMML.XSL        # WPS ships a copy too
```

A typical hit: `C:\Program Files\Microsoft Office\root\Office16\MML2OMML.XSL`.

On Linux or macOS there is no Office, so there is no stylesheet. Copy the file
from a Windows install and set `MML2OMML_XSL` to its path — the stylesheet is
plain XSLT 1.0 and libxslt applies it fine anywhere.

## Font of the equations

Word renders math in Cambria Math by default. This chain injects
`Times New Roman` instead, by giving every `<m:r>` a
`<w:rPr><w:rFonts/></w:rPr>` with `ascii`, `hAnsi`, `cs`, and `eastAsia` all set.

Control it with `MATH_FONT`:

| Setting | Effect |
|---|---|
| unset | Times New Roman |
| `MATH_FONT="Cambria Math"` | Word's default look |
| `MATH_FONT=""` | no injection at all |

Two details worth knowing:

- The **italic** of math variables is controlled by `<m:sty>`, not by the font.
  Changing the font family does not flatten the italic.
- Injection is complete: every `<m:r>` in every generated formula gets the
  `w:rPr/w:rFonts` block, including runs nested inside `m:f`/`m:rad`/`m:nary`
  argument containers. There is no residue on Cambria Math in the visible text.

The count is the cheap way to check a document after conversion. On the smoke
fixture, which is the regression baseline this chain is measured against:

```
_smoke/in.docx          → 4 formulas / 0 failures / TNR=80 / limUpp=0
```

`out_tnr.docx` is that baseline's recorded output. `TNR=80` is the number of
`w:rFonts` attributes set to Times New Roman — 20 `<m:r>` runs times four
attributes each — and it is the number that regresses first if injection ever
stops reaching a nesting level. `limUpp=0` says no phasor dot survived in its
wrong form. The count moves with the content: a document with a phasor that the
patch rewrites legitimately loses runs, because the three runs of an `m:limUpp`
collapse into one `m:acc`. Compare run counts between documents only after
accounting for that.

## What the span detector refuses to convert

`find_math_spans()` is deliberately conservative, because false positives are far
worse than misses: a mangled sentence is a visible defect, while a formula left
as plain text still reads.

It refuses:

| Pattern | Example | Reason |
|---|---|---|
| interval | `[0, 1]` | single line without `;` — not a matrix |
| function call | `f(x, y)` | parentheses + commas, no operator |
| bare quantity | `6V`, `2Ω` | no relation operator |
| fragment | `b=`, `=R` | missing an operand |
| date | `2026-11-08` | digits-and-dashes |
| CJK break | `范围 = 前 5 周` | no right-hand operand |

It accepts `U=IR`, `ΣI=0`, `f=50Hz`, `i_{L}(0^{+}) = i_{L}(0^{-})`,
`A = [a, b; c, d]`.

Accepting a span is not the same as emitting an equation object for it — the tier
split decides that. Of that list, `U=IR`, `ΣI=0`, and the two subscripted forms
stay OMML; `f=50Hz` lands in tier 3 as plain text, because it is a quantity
symbol, one `=`, a number, and a whitelisted unit.

## Marking spans yourself with `$...$`

Everything above assumes the author wrote linear notation and left the detector to
find it. The detector is conservative by design, so it will sometimes decline a
span you know is a formula — and no amount of tuning fixes the general case,
because the information it lacks is *intent*, which only you have.

The escape hatch is to mark the span yourself. `$...$` is not LaTeX in the
document sense and pandoc will not see it; it is a convention this chain honours,
and the rule is one line: whatever sits between the dollars goes to the converter
as LaTeX, verbatim, and the text around it stays text.

```python
def _add_mixed(p, text):
    pos = 0
    for m in re.finditer(r"\$([^$]+)\$", text):
        if m.start() > pos:
            _run(p, text[pos:m.start()])
        tex = m.group(1)
        try:
            el = latex_to_omml_lxml(tex)
            fix_phasor_dots(el)      # 手写调用方必补，否则点会浮大
            p._p.append(el)
        except Exception as exc:     # noqa: BLE001
            FAILS.append((tex, str(exc)[:150]))   # 记账回退，绝不静默丢公式
            _run(p, tex)
        pos = m.end()
    if pos < len(text):
        _run(p, text[pos:])
```

Three things in that loop carry weight:

- the regex is `\$([^$]+)\$`, non-greedy by construction — an empty span matches
  nothing, and an unterminated `$` is left as literal text rather than swallowing
  the rest of the paragraph;
- `fix_phasor_dots(el)` is called by hand, because this loop is itself a
  hand-written caller — nothing below it in the chain will do the repair (see
  "Phasor dots need their own accent");
- the `except` branch does not drop the formula. It records `(latex, error)` in
  `FAILS` and falls back to emitting the source text, so a formula that fails to
  convert shows up as its own LaTeX on the page — visible, greppable, and listed
  in the run summary. Silently losing an equation is the one outcome worth this
  much code to prevent.

The trade-off is honest: a `$` in prose now has to be written `\$`, and the
detector's opinion no longer applies inside the markers. In exchange you get
exactly the equations you meant, in exactly the places you meant them.

## Phasor dots need their own accent

`\dot{X}` looks like one operator, but `latex2mathml` reads it as an accent applied
to a base, and emits the accent through `m:limUpp` — the limits-upper construct,
which is what an over-something operator becomes when its content is not
recognised as a combining mark. Word lays `m:limUpp` out as a stacked limit: the
dot sits above the letter at full size, off-centre, and stays put when you retype
the base. The dot of a phasor is not a limit; it is a diacritic, and Word has an
element for that: `m:acc` with the combining dot above, `\u0307`.

`fix_phasor_dots()` rewrites one into the other. It walks the tree, finds every
`m:limUpp` whose body is a single letter carrying exactly one dot, and replaces
it with `m:acc` + `m:chr[val="\u0307"]`, deep-copying the base run's `w:rPr` into
the new `m:e` so the Times New Roman injection survives the move. It is
deliberately narrow — a single-letter base and a single dot on top. A genuine
`\lim` has a multi-character body and is passed through untouched, so the same
tree can hold both a phasor and a limit.

Measured on three inputs, counting `m:acc` and `m:limUpp` occurrences in the
serialised OMML:

```
'\dot{I}'                    fixed 1   acc=2  limUpp=0
'\dot{U}_{AB}'               fixed 1   acc=2  limUpp=0
'\dot{I}_{L}+\dot{I}_{C}'    fixed 2   acc=4  limUpp=0
```

`acc=2` for a single dot is not a double application: `m:acc` contributes an
opening tag and a closing one, and the counter matches both.

`latex_to_omml_lxml()` does **not** fix this, on purpose. It is a pure converter —
LaTeX in, OMML out, no post-processing — and `\dot{I}` through it still measures
`acc=0 limUpp=1`. The repair is applied by the layer above: in `docx_mathify.py`
every conversion goes through `_omml_with_phasor()`, which is the single exit for
the whole module, and in `insert_formula.py` the fix runs immediately after the
conversion inside `_omml_stdlib()`.

Both CLIs report the count, and a failure to repair never sinks the run:

| Where | What you see |
|---|---|
| `docx_mathify.py` summary line | `修正相量点 N` |
| `docx_mathify.py --report` | a summary row plus, when there were failures, a detail block |
| `docx_mathify.py` stderr | `PHASOR-WARN <latex> -> <ExceptionType>: ...`, one per failure |
| `insert_formula.py` stdout | `修正相量点 N 处`, appended only when N > 0 |

`stats["phasor_fixes"]` is initialised to `0` and `stats["phasor_errors"]` to
`[]`, so a run with no phasors reports a zero rather than omitting the field.
`_omml_with_phasor()` wraps the repair in `try/except`: an exception is recorded
in `phasor_errors` (with the exception type name) and the **unrepaired but usable**
OMML is returned, so a broken accent never drags the whole document down into
`stats["failed"]`. `insert_formula.py` mirrors this with the module-level
`_LAST_PHASOR_FIXES` / `_LAST_PHASOR_ERROR` pair and prints the failure to stderr.

That makes the count a supported output you can read — the summary line and the
report are the contract, so a regression in the repair shows up as a number that
stops moving. Diagnostics go to stderr and never mix into stdout.

If you are a **hand-written caller** — your own script doing
`from math_omml import latex_to_omml_lxml` and appending the result yourself —
call `fix_phasor_dots(el)` before you place it. That return value is the count the
repair hands back, and asserting on the tree afterwards is how you prove it
happened:

```python
el = latex_to_omml_lxml(tex)
fix_phasor_dots(el)
assert '<m:limUpp' not in etree.tostring(el, encoding="unicode")
```

One more line from the same measurement run is worth explaining so it does not
look like a defect. `\dot{I}_{L}+\dot{I}_{C}` written alone, with no relation
sign anywhere in the span, is not converted at all — no OMML, nothing to fix.
That is the span detector behaving as designed: `find_math_spans()` requires a
relation character or one of the strong symbols, and `+` is an operator, not a
relation, so it is not in `OP_CHARS`. Add `= \dot{I}` and the whole thing becomes
one span with two accents.

## Merged scripts

Word does not store "sub 0, then letter L" as two scripts. It stores a run whose
`vertAlign` says subscript and whose text is `0L`. Converting that literally
gives `\omega_{0L}` — one subscript containing two tokens — when the author meant
`\omega_0 L`.

`repair_merged_scripts()` splits these apart heuristically: it recognises when
the tail of a script is actually a following token, and honours the cases where
it genuinely is one script (`U^{2}`, `10^{-6}`).

## Two defects fixed in this round

Both live in `linear_to_latex.py` and both are about source-text shapes that the
converter used to read wrong. Neither touches `docx_mathify.py`'s tier split.

### Consecutive Unicode superscripts were split into separate groups

`_escape_unicode` translated literal Unicode script characters one at a time, so
each character of a multi-character exponent emitted its own script group.
`10⁻⁷` is one exponent written as two characters, and it came out as
`10^{-}^{7}` — an empty superscript group followed by a second superscript. That
is not merely ugly, it is invalid LaTeX: `latex2mathml` keeps the first group and
drops or flattens the second, so the document showed a `10` with an empty
exponent, or a stray body-text `7` beside it. Subscripts had the identical flaw
(`₁₂`), just unreported.

The fix merges each maximal run of same-kind characters into one group before the
per-character loop runs. Only same-kind characters merge, so `x²₁` still yields
`x^{2}_{1}` in that order. Single-character runs produce byte-identical output to
the old code, which is why no existing assertion had to change.

This also repaired three real classes of error found in the sample corpus:
chemical charges (`Fe^{3}^{+}` → `Fe^{3+}`, `SO_{4}^{2}^{-}` → `SO_{4}^{2-}`),
negative unit exponents (`L^{-}^{1}` → `L^{-1}`), and scientific notation
(`4\times 10^{-}^{7} m` → `4\times 10^{-7} m`).

### A unit after an exponent was read as a fraction

The dimension-span finder runs before Unicode escapes are applied, so it sees the
source shape. For `c=3×10⁸m/s` the `⁸` was still a literal character, and it
satisfied neither anchor the finder accepts: it is not one of the ASCII forms the
anchor regex knew (a digit, a `^{n}` closing brace, a closing parenthesis), and
the left-boundary check rejected it because `⁸`.isalnum() is true. With no valid
anchor the span went to the fraction builder and produced
`\frac{10^{8}m}{s}`.

The anchor regex now also accepts runs of literal superscript *digits*, so
`c=3×10⁸m/s` becomes `c=3\times 10^{8}\,\text{m/s}`. The character class holds
digits only: `⁺` and `⁻` are signs with no value to anchor on, and `ⁿ` / `ⁱ` are
letters, so `xⁿm/s` still reads as a fraction.

The governing principle is parity, not permissiveness. The braced spelling
`c=3×10^{8}m/s` already produced a dimension string before this change; only the
literal spelling diverged. Both spellings now give the same result, which is the
point — two spellings of one fact producing two different outputs was the bug.
The same test already existed for `kg/m3` versus `kg/m³`.

## Escaping traps in source text

`docx_mathify.py` reads whatever linear notation is in the document. A few
characters break the pipeline if they reach the converter unescaped. Fix them
while writing the source text, not afterwards:

| Character | Problem | Write instead |
|---|---|---|
| `%` | LaTeX comment character | `\%` |
| `℃` | not in the math font | `°C` |
| `～` (full-width tilde) | not a math symbol | `~` or `\sim` |
| `\|x\|` | bare pipes | `\left\|x\right\|` |
| `[]²` | split across a script boundary | keep the bracket and exponent together |
| compound units | `N·m` read as a product | `\text{N·m}` |

`--dry-run` reports conversions without writing, which is the cheap way to audit
a large document before committing to the change.

## Command line

Two entry points. `docx_mathify.py` handles a whole document; `insert_formula.py`
handles one equation at a known anchor.

```
python scripts/docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex
python scripts/insert_formula.py unpacked/ "E=mc^2" --after "marker"
```

Both run from the repository root — the `scripts/` prefix is part of the command.
Dropping it (`python docx_mathify.py …`) fails with `can't open file …docx_mathify.py`
and exit `1`, because the script is not in the working directory.

### `docx_mathify.py`

| Flag | Effect |
|---|---|
| `input` | the source `.docx`, opened read-only |
| `-o`, `--out` | where to write the converted document |
| `--inplace` | overwrite the source, after copying it to `<name>.bak.docx` |
| `--report` | write a Markdown table of original notation ↔ generated LaTeX |
| `--tex` | write a deduplicated list of the LaTeX the run produced |
| `--no-merged-script-fix` | leave merged sub/superscripts alone (`\omega_{0L}` stays one script) |
| `--dry-run` | report only; no document is written |
| `--number-align` | number the rows of an `align`/`gather` that carry no `\tag` |

At least one landing flag is required, and the three flags are **not** three
interchangeable choices: each pair is a usage error. `-o` and `--inplace` name two
different write targets; `--dry-run` writes nothing at all, so a path given to `-o`
can never come into existence. Measured 2026-11-11 on the same input document
(SHA256 checked before and after each run):

| Flags | Exit | What is written |
|---|---|---|
| `-o out.docx` | `0` | `out.docx`; the source is untouched; no `.bak` |
| `--inplace` | `0` | the source is rewritten after a copy to `<name>.bak.docx` |
| `--dry-run` | `0` | nothing at all |
| `-o out.docx --dry-run` | **`2`** | nothing; `out.docx` is not created, stdout is empty |
| `--inplace --dry-run` | **`2`** | nothing, and **no** `.bak` (argparse rejects the pair before any work) |
| `--inplace -o out.docx` | **`2`** | nothing; neither file is touched, no `.bak` |
| none of the three | **`2`** | nothing |

Both refusals are **usage errors, not silent choices**: writing to one path while
overwriting another is a contradiction, and `--dry-run` announcing a destination it
will never produce is a lie the caller cannot detect. Each is rejected **before any
read or write**, so the input is byte-identical afterwards. That is deliberate — a
batch rewrite that quietly does nothing, or one that picks a destination you did not
mean, looks identical to success until someone opens the document expecting
formulas.

### Exit codes, and `3` in particular

| Code | Meaning |
|---|---|
| `0` | the document was written (or `--dry-run` wrote nothing, as asked) and no span failed |
| `1` | expected failure that is **not** about the input path — the conversion chain raised on a document it did manage to open |
| `2` | usage error, decided **before** any read or write: `--inplace` with `-o`, `--dry-run` with either of them, or none of the three landing flags — and an input path that is missing, is a directory, or is not a Word package |
| `3` | **partial success**: the document was written, but at least one formula did not convert and that span stays as its source text |

A missing input file used to sit in row `1`. It belongs in row `2`, because the
path is checked before anything is read: nothing is written either way, but a
caller that branches on the code wants "your argument was wrong, fix it and retry"
(`2`) to be distinguishable from "the document itself has a problem" (`1`). The
`--dry-run` / `-o` conflicts are checked first, so a run with both a bad flag pair
*and* a missing path reports the flags.

`3` is the only exit that combines an artifact with a non-zero code, so a caller
must check it — a silent `0` would let a document with broken equations ship as
finished. It is emitted when any span reaches the converter and fails: a missing
or unparseable `MML2OMML.XSL`, a `latex2mathml` rejection, an XSLT failure, or
`assert_legal()` catching an illegal OMML shape.

A `latex2mathml` rejection is the cheapest way to see it, and needs no broken
environment — write the span `F=x/y` with `%` as a real superscript run, so the
linear form reads `F=x/y%`:

```python
# spec for docx_create: the % must be a superscript run, not a literal caret
{"type": "paragraph", "runs": [{"text": "F=x/y"},
                               {"text": "%", "vert_align": "superscript"}]}
```

```
python scripts/docx_mathify.py pct.docx -o pct-out.docx
```

Measured: exit `3`; stderr `error: 1 formula(s) failed to convert` plus
`FAIL F=\frac{x}{y^{%}} -> latex2mathml 转换失败`; and `pct-out.docx` **is still
written** with that span left as source text. `%` is LaTeX's comment character, so
the group never closes and the converter refuses the whole span — which is exactly
the partial-success shape `3` exists to report.

Missing `MML2OMML.XSL` is not a synonym for `3`: a document whose spans all
degrade to plain text runs never touches the stylesheet and still exits `0`; only
a span that needs an equation object (a fraction, a radical, …) reaches the
failure. `3` belongs to `docx_mathify.py` — and to `docx_template.py apply
--fix-formulas`, which passes mathify's code back up unchanged (`1`/`2` as-is, `3`
as `3`, and only `0` collapses to `0`), because it runs mathify in place. No other
script in `scripts/` returns it.

`--report` and `--tex` are worth running together. The report is for reading: it
pairs each original notation with the LaTeX that came out, one row per span,
which is how you spot a span that was converted but converted wrongly. The `.tex`
file is for diffing: it is a deduplicated list, so on a re-run a change in it means
the converter's output changed, not that the document did.

### `insert_formula.py`

The first positional argument is `unpacked_dir` — a directory you have already
extracted from the `.docx`, not the `.docx` itself, because the script edits
`word/document.xml` in place and expects to hand the tree back to a packer.
The second is the LaTeX.

| Flag | Effect |
|---|---|
| `--after MARKER` | place the equation after the paragraph containing that text |
| `--before MARKER` | place it before that paragraph |
| `--display` | emit `m:oMathPara` — a centred display equation rather than inline |
| `--raw-spacing` | keep `\quad`, `\,`, `\;`, `\:` as written |
| `--dump` | print the generated OMML and write nothing |

`--after` and `--before` are mutually exclusive, and a marker that is not found is
an error (`exit 1`), not a fallback to the end of the document. A formula in the
wrong place is harder to notice than a command that refused.

The other three refusals are usage errors (`exit 2`, decided before anything is
read or written), because each one used to be a silent swallow:

- `--dump` together with `--after`/`--before` — `--dump` only prints the OMML, so
  the anchor would be dropped on the floor;
- `--after` together with `--before`;
- an **empty** marker (`--after ""`) — it would append the formula at the end of
  the document instead of anchoring it.

On success the stdout carries the anchor the script *measured* after inserting,
not the argument echoed back: `Inserted inline formula: E=mc^2  at body[1]`
(`at body[0]`-style indices are the container's own child order, and `at tc[i]`
appears when the marker sits inside a table cell). A marker that is missing
reports the searchable range instead of staying vague:
`error: marker not found: 'x' — searched 3 body paragraph(s) (0-based [0, 2]) and 1 table cell(s)`.

Two things the script does for you:

- An inline equation that would otherwise be the only content of its paragraph
  gets a zero-width space run appended (`\u200b`). Without it Word promotes the
  `m:oMath` to a display equation and the layout changes under you — the same
  problem pandoc records as issue #11674.
- Spacing commands are rewritten before conversion. `\quad`, `\,`, `\;`, `\:`
  and `\!` do not survive the trip through MathML and the stylesheet; the
  characters are dropped and the formula closes up. `keep_spacing()` rewrites each
  of them to `\ `, a control space, which does survive. `--raw-spacing` turns the
  rewrite off, and is only useful when you are inspecting what the XSLT drops.
