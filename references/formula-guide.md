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
