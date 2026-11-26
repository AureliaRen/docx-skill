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

### Two layers, and only one of them is the body

A document carries two independent font layers. Changing one does not move the
other, and it is easy to spend an afternoon on the wrong one:

| Layer | Where it lives | Who writes it |
|---|---|---|
| body text | `<w:rPr><w:rFonts>` on an ordinary `<w:r>` | the styles, `docx_create`, `fix_fonts.py` |
| equation text | `<w:rPr><w:rFonts>` **inside** each `<m:r>` | `math_omml.apply_math_font()`, from `MATH_FONT` |

Both layers get the same four attributes — `ascii`, `hAnsi`, `cs`, `eastAsia` —
but from different code, at different times. The rule that follows is narrow and
absolute:

> **Switching the body's CJK font must not touch the formula's `eastAsia`.**

`\text{中文}` is an equation run, so it takes the *formula* layer's `eastAsia`,
not the body's. Rewrite the body to 楷体 and leave the formula's `eastAsia` behind
and the Chinese inside `\text{}` no longer matches the surrounding prose — while
every Latin glyph still lines up, so the document looks correct until someone
compares the two lines. Change the formula font **as a formula font**, through
`MATH_FONT`:

```
MATH_FONT="楷体"   python scripts/docx_mathify.py in.docx -o out.docx
```

`fix_fonts.py` already obeys this split: for a run inside `m:oMath` it rewrites
`ascii`/`hAnsi`/`cs` to Times New Roman and deliberately leaves `eastAsia` alone
(`scripts/fix_fonts.py:13-14`, implemented at `scripts/fix_fonts.py:179-190`, where
the math runs are collected by `id()` before the rewrite loop). Do not "finish the
job" by adding `insert(0, rFonts)` for every run — see
[font-scheme](font-scheme.md#和公式链的交界) for why that breaks `m:r` child order.

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
document sense and pandoc will not see it; it is a convention the loop below
implements, and the rule is one line: whatever sits between the dollars goes to
the converter as LaTeX, verbatim, and the text around it stays text.

Read that as **a recipe for code you write**, not a description of the shipped
CLIs. `docx_mathify` and `docx_create` have no marker layer: a `$` in a paragraph
is an ordinary character to them, so `$x=1$` survives as literal text with
`oMath=0` and nothing is converted. The loop is what you add when your own script
needs the marker, and it is the only place the delimiters are consumed.

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

The trade-off is honest: **inside your own loop**, a `$` in prose has to be written
`\$`, and the detector's opinion no longer applies between the markers. In
exchange you get exactly the equations you meant, in exactly the places you meant
them. (In the shipped CLIs there is nothing to escape — see the currency note at
the end of the next section.)

### Doubled dollars are not a marker

The marker is **single** `$...$`. `$$...$$` is not a display marker, not a
stronger marker, and not recognised at all: the convention is one dollar on each
side and the pattern in the recipe above — `\$([^$]+)\$` — has no second-dollar
branch. Anyone who writes LaTeX habitually reaches for the doubled form, which
makes this the trap that gets hit most often, and it fails in the worst possible
way.

| Text in the source | Visible text after conversion | `oMath` |
|---|---|---|
| `由 $\sigma=F/A$ 可知，且 $\sigma=\frac{F}{A}$ (3-1)` | `由 $$ 可知，且 $$ (3-1)` | 2 |
| `由 $$\sigma=\frac{F}{A}$$ 可知` | `由 $$$$ 可知` | 1 |
| `$$\sigma=\frac{F}{A}$$ (3-1)` | `$$$$ (3-1)` | 1 |
| `σ = F/A` (bare linear form, no markers) | (a zero-width space only) | 1 |

The mechanism, measured end to end (`docx_create` → `docx_mathify` → read back
`w:t`): **the detector finds the LaTeX on its own and removes it from the text
runs; every `$` the author typed stays behind.** In row 2 the span
`\sigma=\frac{F}{A}` is lifted out and converted, leaving the two dollars from
each side adjacent — four characters that now read as `$$$$`. In row 1 the same
thing happens twice, so you get `$$` in two places. Nothing in the chain ever
looked for a closing dollar; the doubled form is not rejected, it is simply not
*seen*, and the residue is what remains after the formula leaves.

That matters for where the fix belongs. The `$...$` loop shown above is a recipe
for your own caller — it is not something the shipped CLIs run. `docx_mathify`
and `docx_create` treat `$` as an ordinary character: `$x=1$` in a paragraph
survives untouched (`oMath=0`), and whenever the detector *does* convert a span
between markers, the markers are left in the prose. So there is no configuration
that makes `$$` work, and nothing to switch off — the correct move is not to
write it.

A run reports `失败 0` and exits `0` while doing all of this, and the equation
itself is correct. Only the page text is wrong.

That combination — silent, exit `0`, formula present — is why this needs a
read-back rather than a summary line. **Grep the converted `w:t` for a residual
dollar:**

```python
import sys, zipfile
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
root = etree.fromstring(zipfile.ZipFile(sys.argv[1]).read("word/document.xml"))
for i, p in enumerate(root.iter(W + "p")):
    text = "".join(t.text or "" for t in p.iter(W + "t"))
    if "$" in text:
        print(f"paragraph {i}: {text!r}")
```

Nothing printed means no residual marker. `docx_read.py --text` shows the same
text and is enough for a single file; the loop above is for sweeping a batch.

Two correct ways to get a centred display equation instead:

- **write a `docx_create` `latex` block** with `"display": true` — it produces an
  `m:oMathPara` and never touches any marker layer;
- **use `insert_formula.py --display`**, which builds the `m:oMathPara` in place.

Note what does *not* work: dropping the markers and letting the detector see the
span gives you an **inline** equation, not a display one. `docx_mathify` emits
`m:oMath` on every path and never an `m:oMathPara` — measured on a paragraph
holding nothing but `\sigma=\frac{F}{A}`, the result is `oMath=1, oMathPara=0`.
The markers exist to rescue spans the detector *declines*; they are not, and never
were, a display-layout request.

A currency sign in prose needs neither form. Measured on `docx_mathify` output:
`单价 $5 一件` comes back untouched (`oMath=0`) — a lone `$` that does not enclose
something the detector reads as a formula is just a character. What you should
*not* do is escape it: `单价 \$5 一件` comes back with the backslash still on the
page, because nothing in the shipped chain maps `\$` back to `$` — that mapping
exists only inside a formula. The escape is for the hand-written `$...$` caller,
whose loop consumes the delimiters; the CLIs do not consume anything.

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

## Function names and differentials come out upright

Word sets letters in a math zone *italic*, because a single letter is a variable.
Digits, operators, and function names are not variables and must stand upright.
`latex2mathml` supplies `m:sty val="p"` in a few places (`\sin`, `\log`,
`\mathrm{...}`) and Microsoft's stylesheet does not add any more, so
`\max\{x,y\}` used to arrive as one run holding the literal text `max{x,y}`,
italic in Word.

Since this round the converter repairs it in `split_runs_for_upright()`
(`scripts/math_omml.py:192`, called from `latex_to_omml_lxml` at
`scripts/math_omml.py:498`): it classifies each run's characters and rebuilds the
run as segments, giving every digit run, operator, and whitelisted function name
its own `m:rPr/m:sty val="p"`.

**Splitting is the mechanism, not an optimisation.** `m:sty` is a *run-level*
property, and one run routinely holds both kinds of character at once:
`\frac{dy}{dx}` produced the numerator as a single run `dy` and the denominator
as a single run `dx`. Setting `sty=p` on `dx` uprights the `d` **and the `x`**,
which is more wrong than leaving both italic. The only correct move is to cut the
run by character class first, so every segment carries its own answer. If you
write your own post-processing pass, it has to split `m:r` elements — a loop that
only adds an attribute to existing runs cannot express this.

### `\mathrm{max}` is still the recommendation, but no longer the only correct form

Measured on this repository, before and after the pass landed:

| Input | Runs produced (before) | Runs produced (now) |
|---|---|---|
| `\max\{x,y\}` | one run, text `max{x,y}`, no style | `max`(p) · `{`(p) · `x`(–) · `,`(p) · `y`(–) · `}`(p) |
| `\mathrm{max}\{x,y\}` | `max`(p) + `{x,y}` | `max`(p) · `{`(p) · `x`(–) · `,`(p) · `y`(–) · `}`(p) |
| `\text{max}\{x,y\}` | `max`(`m:nor`) + `{x,y}` | `max`(`m:nor`) · `{`(p) · `x`(–) · `,`(p) · `y`(–) · `}`(p) |

The first two rows are now **segment-for-segment identical**, and Word renders them
identically: exported through `pdf_export.ps1` and compared pixel by pixel, the two
`max` regions differ in **0 pixels**. So `\mathrm{max}` stays the written
convention because it says what it means to every other renderer and to a human
reading the `.tex`, but `\max\{x,y\}` is no longer wrong.

Before the pass they did *not* agree, and the size of the disagreement is a
property of the rasteriser rather than of the document — quote it with the DPI or
not at all:

| Rasterisation DPI | 72 | 96 | 150 | 200 | 300 |
|---|---|---|---|---|---|
| differing pixels, old vs new | 371 | 556 | 1195 | 1957 | 3896 |

(sample: one `latex` block reading `\max\{x,y\}\le z`, against the same block
with `\mathrm{max}`; `pdf_export.ps1` → PDF, `fitz` rasterises at the stated DPI,
`ImageChops.difference`, non-zero pixel count on a 1275x1650 page.)

The third row is worth avoiding for a different reason: `\text{}` forces
`m:nor`, the body-text style, so the glyph is upright by accident rather than
because it is an operator — and everything else in the group inherits text
spacing with it.

Two scope limits, both deliberate:

- the whitelist (`UPRIGHT_FUNCS`, `scripts/math_omml.py:112`) holds 34 names and
  **excludes `E`**: in engineering and physics `E` is nearly always EMF or energy,
  an italic variable, so uprighting it wholesale would damage far more formulas
  than the rare expectation operator `E[·]` would gain;
- a run that already carries `m:nor` or a non-empty `m:sty` is skipped whole, so
  `\sin`, `\mathrm{Var}` and `\text{...}` keep whatever the upstream decided.

## `\begin{cases}` and matrices are the same object in OMML

`cases`, `pmatrix`, `bmatrix`, `vmatrix`, `matrix` and `array{ll}` all convert to
the *same* OMML shape: an `m:m` whose `mPr` carries `baseJc=center` and
`plcHide=on`, with `mcs/mc/mcPr/{count, mcJc=center}`. Measured across all six
environments the `mPr` is identical. Nothing in the OMML says which one it was.

The distinction survives in exactly two places, and both are upstream of the
stylesheet:

- **the LaTeX source** — `\begin{cases}` versus `\begin{pmatrix}`;
- **the MathML layer** — every `mtd` of a `cases` carries
  `columnalign="left"`; a matrix's `mtd` carries no `columnalign` at all
  (`array{ll}` carries `left` too, and that is correct — it asked for `l`).

`align_cases_left()` (`scripts/math_omml.py:323`) therefore reads the MathML tree
that `latex_to_omml_lxml` still has in hand, decides which tables are cases-like,
and only then rewrites the matching `m:m` to `mcJc=left`.

**Why a blanket pass over `mcJc` is not an option.** "Every `mcJc` becomes
`left`" is one line, and it silently left-aligns every matrix in the document as
well. Because the mapping is positional, the count has to agree first: only tables
with **two or more columns** produce an `m:m`, and they map 1:1 in document order.
When the counts disagree the function changes **nothing** and records the reason in
`stats["cases_align_skipped"]` — a guessed off-by-one would move the left alignment
onto a matrix, and no test would catch it.

Measured, after the pass: `cases` → `mcJc=left`; `pmatrix`, `bmatrix`, `vmatrix`,
`matrix` → still `center`; `array{ll}` → `left`; `cases` nested inside a
`pmatrix` → the inner one goes `left`, the outer stays `center`.

One boundary worth writing down, because it breaks the simple mental model: the
mapping is **not** strictly 1:1 with MathML `mtable` elements. A `pmatrix` wrapping
a `cases` produced two `mtable`s in MathML but only one `m:m` in OMML — the outer
matrix was folded into `m:d` + `m:eqArr`. Compare counts of "tables with two or
more columns", not of all tables.

## A unary sign at the head of a delimiter group

A `\left…\right` pair, and a plain bracketed group too, are converted by the
stylesheet into `m:d` — a delimiter object whose content is a sequence of `m:e`
slots, with `m:sepChr` naming the character that separates them. Those slots are
built by reading operators as separators, and that rule has one failure mode worth
knowing by heart: **a group whose first item is a unary sign, with no other
operator behind it, gets an empty first slot.**

```
\left(-\frac{y}{x}\right)
→ <m:d><m:dPr><m:sepChr m:val="−"/></m:dPr><m:e/><m:e><m:f>…</m:f></m:e></m:d>
                                  ^^^^^^ empty slot
```

Word renders that empty `m:e` as an empty box **inside** the brackets. Everything
else about the equation is correct — the fences stretch, the fraction is a real
fraction, the text layer survives — and that is exactly why it ships. Through
`pdf_export.ps1` the box is in the PDF as well: it is OMML content, not a
Word-only artefact.

### Why every existing check stayed green

| Layer | What it reports | Why that is not enough |
|---|---|---|
| the conversion call | returns an element | an empty `m:e` is schema-legal OMML |
| `docx_mathify` summary | `失败 0`, exit `0` | nothing failed; the formula *converted* |
| `assert_legal()` before this round | passes | it checked `m:oMath` nesting and `m:rad` only |
| `docx_validate` before this round | `{"ok": true, "issues": []}` | it carried no OMML check at all |
| `docx_preview` / LibreOffice | a rendered page | an empty slot is nearly invisible at preview DPI |

The only reliable witness is the OMML itself: count `m:d` elements that own an
empty `m:e`. See [Checking a document for empty slots](#checking-a-document-for-empty-slots).

### The trigger, measured

Run through the 1.3.0 pipeline (`latex_to_omml_safe`, then count `m:d`, empty
`m:e` children and any `m:sepChr` value in the serialised OMML):

| LaTeX | `m:d` | empty `m:e` | `sepChr` |
|---|---|---|---|
| `\left(x\right)` | 1 | 0 | `''` |
| `\left(x+y\right)` | 1 | 0 | `'+'` |
| `\left(\frac{a}{b}\right)` | 1 | 0 | `''` |
| `\left(-x\right)` | 1 | **1** | `'−'` |
| `\left(-x y\right)` | 1 | **1** | `'−'` |
| `\left(-1\right)` | 1 | **1** | `'−'` |
| `\left(-\frac{y}{x}\right)` | 1 | **1** | `'−'` |
| `\left(+\frac{a}{b}\right)` | 1 | **1** | `'+'` |
| `\left(\mp x\right)` | 1 | **1** | `'∓'` |
| `\left[-x\right]` | 1 | **1** | `'−'` |
| `\left\{-x\right\}` | 1 | **1** | `'−'` |
| `\left(\right)` | 1 | **1** | `''` |
| `\left(-x+y\right)` | 1 | 0 | `'+'` |
| `\left(-1+2\right)` | 1 | 0 | `''` |
| `\left(-\infty\right)` | 1 | 0 | `''` |
| `\left(\pm x\right)` | 1 | 0 | `''` |
| `\left[-(a+b)\right]` | 1 | 0 | `''` |
| `\left(\frac{-y}{x}\right)` | 1 | 0 | `''` |
| `\left.\frac{a}{b}` + right-dot | 0 | 0 | — |
| `\left\langle -x\right\rangle` | 0 | 0 | — |

Two rows carry the corrections that matter more than the table itself.

**It is not a round-bracket problem.** Change the fence and nothing changes:
square brackets and braces break the same way, because the stylesheet keys on the
*operator*, not on the fence character. A report that clears `[…]` and `|…|` of
blame is wrong on that point.

**It is not a `\left`/`\right` problem either.** The `fence="true"` attribute
`\left…\right` puts on its `<mo>` is never read by the stylesheet; a bare
`(-\frac{y}{x})` becomes the same `m:d` with the same empty slot. So there is no
bracket style to switch to: "just write ordinary parentheses" was never an escape —
both spellings take this path, and both are cleared by the one MathML pass described
below. What the bare forms do document is the reach of the stylesheet rule:

| LaTeX | `m:d` | empty `m:e` | `sepChr` |
|---|---|---|---|
| `(-\frac{y}{x})` | 1 | **1** | `'−'` |
| `(-x)` | 1 | **1** | `'−'` |
| `(-1)` | 1 | **1** | `'−'` |
| `(-x y)` | 1 | **1** | `'−'` |
| `(-\sin x)` | 1 | **1** | `'−'` |
| `(-\sqrt{2})` | 1 | **1** | `'−'` |
| `(-a^{2})` | 1 | **1** | `'−'` |
| `[-x]` | 1 | **1** | `'−'` |
| `\{-x\}` | 1 | **1** | `'−'` |
| `(-\frac{a}{b}+1)` | 1 | 0 | `'+'` |
| `(-x,\ y)` | 1 | 0 | `','` |
| `(-x)(-y)` | 1 | 0 | `''` |
| `(-x)-(-y)` | 1 | 0 | `''` |

Absolute-value groups belong to the same family. A pipe character splits a
Markdown cell, so the pipe rows below are escaped the way the rest of this guide
does it (`\|`), and two rows use the `\vert` / `\Vert` aliases to show they
behave identically. Every one of the double-bar spellings measured
`['', 'x']` with `sepChr='−'` before the pass and `['−x']` with an empty
`sepChr` after it — the alias changes nothing.

| LaTeX | `m:d` | empty `m:e` | `sepChr` |
|---|---|---|---|
| `\left\| -x \right\|` | 1 | **1** | `'−'` |
| `\|-x\|` (no `\left`/`\right`) | 1 | **1** | `'−'` |
| `\left\vert -x \right\vert` | 1 | **1** | `'−'` |
| `\left\Vert -x\right\Vert` | 1 | **1** | `'−'` |
| `\vert -x \vert` (no `\left`/`\right`) | 0 | 0 | — |

The condition is narrower than "a leading sign", and the reason is mechanical:
the stylesheet takes the *first operator it meets* as the slot separator. Put a
`+` or a `,` behind the sign and **that** becomes the separator, leaving the minus
where it belongs, inside the first slot. Every `sepChr` cell above is a **post-pass**
reading, which matters for the clean rows: `\left(-x+y\right)`,
`(-\frac{a}{b}+1)` and `(-x,\ y)` had no empty slot before the pass either, but
each was a *single* slot holding the whole body with an empty `sepChr`; the pass
turns them into two slots separated by the group's real operator. Rendering is
unaffected — an empty separator is not a rendering feature. Adjacent groups cancel
out for the same reason: in `(-x)(-y)` the second group's opener is not an
operator, so the whole body lands in one slot (measured: `sepChr=''` before and
after, zero empty slots).

### The cause is in `MML2OMML.XSL`, not in this skill

This is the part worth internalising, because it decides where a fix belongs.
Feed the MathML that `latex2mathml` produces straight into the stylesheet — no
skill code anywhere in the path — and the broken `m:d` comes out unchanged:

```python
from lxml import etree
import latex2mathml.converter as lconv
from math_omml import xslt            # cached (transform, path) pair

mml = etree.fromstring(lconv.convert(r'\left(-\frac{y}{x}\right)').encode('utf-8'))
out = xslt()[0](mml).getroot()
s = etree.tostring(out, encoding='unicode')
print(s.count('<m:e/>'), 'empty slot(s)')     # 1 empty slot(s)
print('sepChr' in s)                           # True, m:val is U+2212
```

If that snippet reproduces the empty slot, no Python in this repository is
responsible, and no post-processing *after* the transform can claim the bug was
never there. The stylesheet ships with Office; it is not patched here.

### The fix: wrap the sign and its operand in the MathML tree

`wrap_leading_unary_in_delim()` (`scripts/math_omml.py`, called from
`latex_to_omml_lxml()` before the transform) walks the MathML tree. For a group
whose first item is `−` / ASCII `-` / `+` / `±` / `∓` and which holds at least one
more item, it wraps that sign plus the next sibling into a fresh `mrow`:

```
before:  ( − frac(y,x) )        three items; the sign reads as a separator
after:   ( [− frac(y,x)] )      two items; the sign is inside a group
```

The stylesheet now sees one item where it saw two, emits no separator, and a
group that was broken becomes a single non-empty `m:e`. Because the pass runs upstream,
nothing downstream rebuilds runs, re-copies a `w:rPr`, or re-runs the upright
split: `split_runs_for_upright` and `align_cases_left` receive the OMML they always
received. The count accumulates in `stats["delim_wraps"]`, next to
`upright_split` and `cases_aligned`.

Measured, byte-for-byte on the serialised OMML:

| LaTeX | before | after |
|---|---|---|
| `\left(-x\right)` | 1 empty slot | **0** |
| `\left(-\frac{y}{x}\right)` | 1 | **0** |
| `\left(+\frac{a}{b}\right)` | 1 | **0** |
| `\left(-1\right)^{n}` | 1 | **0** |
| `\left[-x\right]` | 1 | **0** |
| `\left\| -x \right\|` | 1 | **0** |
| `\left\{-\frac{y}{x}\right\}` | 1 | **0** |
| `\left(-\frac{1}{2}\right)\left(+\frac{a}{b}\right)` | 2 | **0** |
| `(-\frac{y}{x})` | 1 | **0** |
| `\left(a+b\right)` | 0, `sep='+'` | 0, `sep='+'` unchanged |
| `\left(a,\ b\right)` | 0, `sep=','` | 0, `sep=','` unchanged |
| `\left(-x+y\right)` | 0 empty, **1 slot** holding `−x+y`, `sep=''` | 0, **2 slots** `['−x','y']`, `sep='+'` |
| `(-\frac{a}{b}+1)` | 0 empty, 1 slot, `sep=''` | 0, 2 slots `['−a/b','1']`, `sep='+'` |
| `(-x,\ y)` | 0 empty, 1 slot, `sep=''` | 0, 2 slots `['−x',' y']`, `sep=','` |
| `\left(\right)` | 1 | **1** — nothing inside to group, and `assert_legal` now rejects it |

Two things to read off that table carefully.

**A clean row is not necessarily an untouched row.** The three `sep` rows were
never broken — zero empty slots — but the pass still rewraps them, because their
first item *is* a unary sign. What moves is structure only: one slot holding the
whole body with an empty `sepChr` becomes two slots separated by the group's real
operator. Word draws `(−x+y)`, `(−a/b+1)` and `(−x, y)` identically either way;
an empty `sepChr` is not a rendering feature, just what the transform defaults to
when no slot boundary exists yet. `\left(a+b\right)` and `\left(a,\ b\right)` open
with an ordinary symbol, so the pass never sees them: byte-identical before and
after, `sep` included.

**The last row is the honest boundary.** A genuinely empty `\left(\right)` has no
second item to pair the sign with, and no unary sign to pair either; the slot rule
still empties it. The pass skips it and the checks below report it.

### Checking a document for empty slots

An older copy cannot repair, but it can still be audited. Open
`word/document.xml` and count the `m:d` elements that own an empty `m:e`:

```python
import zipfile
from lxml import etree

M = "http://schemas.openxmlformats.org/officeDocument/2006/math"

def empty_slots(path):
    with zipfile.ZipFile(path) as z:
        root = etree.fromstring(z.read("word/document.xml"))
    return [d for d in root.iter("{%s}d" % M)
            if any(len(e) == 0 for e in d.findall("{%s}e" % M))]

for delim in empty_slots("report.docx"):
    print(etree.tostring(delim, encoding="unicode")[:160])
```

A count of `0` is the clean answer. Anything else names the group to fix. Note
that *empty* here means structurally empty (`<m:e/>`, no children): a slot holding
a single space is not what Word draws as a box.

Two built-in readers now report the same condition:

- **`assert_legal()`** raises `RuntimeError` when an `m:d` holds an empty `m:e`, as
  a third structural check beside the `m:oMath`-nesting and `m:rad` cases. Every
  conversion path runs through it, so a formula cannot reach a document unnoticed.
  The message splits the diagnosis in two, and neither branch is "write it with a
  different bracket". **The sign-at-head shape** (`wrap_leading_unary_in_delim`'s
  own case): on the normal path it is already repaired, so seeing it means the
  formula bypassed `math_omml` — hand-built OMML from somewhere else, which no
  MathML pass could have touched. **A genuinely empty group** (`()` with nothing
  between the delimiters, or `\left(\right)`): there is no content to place, and the
  toolkit does not invent any. Both branches end in the same advice — complete the
  formula, and if it still trips, check whether it went around `math_omml`.
- **`docx_validate`** reports it as an `omml_empty_slot` issue. Reported, never
  fatal: `ok` and the exit code are unchanged, and a document with no equations
  never grows the key.

### Escaping it on an older copy

On a copy that predates the pass, in order of preference:

1. **Upgrade.** The pass runs inside `latex_to_omml_lxml()`, so every entry point
   inherits it — `docx_mathify`, `insert_formula`, and `docx_create` `latex`
   blocks alike. No spec changes.
2. **Audit what you already shipped** with the snippet above and re-convert only
   the offending paragraphs. Re-saving a document does not heal it: the empty
   `m:e` survives every edit of the surrounding prose.
3. **Do not rewrite your way out of it, and on an upgraded copy there is nothing
   to rewrite.** Ordinary parentheses take the same stylesheet path (`(-\frac{y}{x})`
   measured one empty slot too), so changing bracket style changes nothing; braces or
   a hand-wrapped sign are just as pointless, since the pass already does that
   regrouping. A literal `\vert` pair is not a substitute for the group you wrote
   either — `\vert -x \vert` happens to produce no `m:d` at all, so it changes
   the rendering (no stretchy fences) rather than fixing anything. Upgrade, and both
   `\left…\right` and bare groups come out with zero empty slots.

None of these is a substitute for upgrading: they narrow the damage, they do not
remove the rule.


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

### A vulgar fraction is a value, and a fraction character is evidence

`½` is one code point that fonts draw as a small `1`, a slash and a lowered `2`.
Left alone it survives a conversion as a raw glyph: it looks *nearly* right in
the font it was written in and wrong in every other, so nothing reports it and
the problem shows up at print. All nineteen Unicode fraction code points
(U+00BC–U+00BE, U+2150–U+215E and U+2189) now translate to `\frac{n}{d}`, and
the full table lives in `linear_to_latex.VULGAR_FRAC`.

A fraction character also **opens a span by itself**, on the same footing as
`√` / `∑` / `∫` and as an explicit `_{...}` group. Two reasons it gets that
standing rather than being made to earn it with an `=` somewhere: it is a math
glyph the author typed on purpose, and its untouched form is the worst possible
rendering. Before this, `x=½` converted while `½+1` stayed as text — the same
character with two outcomes depending on whether a relation operator happened to
be nearby. It is also exempt from the `len >= 2` floor, so a lone `½` between
Chinese characters converts while the framing stays text.

And it is a **value anchor** for the unit grammar, which is what keeps the two
meanings apart:

```
½ m/s     → \frac{1}{2}\,\text{m/s}     measurement: half a metre per second
0.5 m/s   → 0.5\,\text{m/s}             the same quantity, same layout
v=m/s     → v=\frac{m}{s}               a formula: m divided by s
m/s       → \,\text{m/s}                a bare unit string
```

The rule behind that column is an anchor check, not "both sides look like
units" — `m/s` and `v=m/s` are identical on that test. Without the fraction
character in the anchor regex, `½ m/s` fell through to the fraction builder and
came out as `\frac{1}{2} \frac{m}{s}`, i.e. "half divided by metres per second",
while `0.5 m/s` on the next line took the unit path. Two adjacent lines of one
hand-out disagreeing is harder to spot than either being wrong.

### A spacing command is one token, not a backslash plus punctuation

`linear_to_latex` splits its input into atoms before any rewriting. `_CMD_RE`
matches `\[A-Za-z]+`, so `\,` — two characters, no letter — never matched, and
`_split_atoms` fell through to its "backslash is a single atom" branch. The comma
then became an *operator* and the bare backslash an *adjacent factor*:

```
1/2 m/s  → 1/2\,\text{m/s} → \frac{1}{2\},\text{m/s}   ← denominator is 2\}
```

The escaped brace closed the group early: OMML came out with numerator `1`,
denominator `2},`, then `m/s`. Word draws a stacked fraction whose denominator
reads `2},`. It opens fine and looks like a fraction from across the desk, so
nobody reported it — the shape that triggers it is "a fraction followed by a
unit string", which is most of a physics hand-out. All five two-character
spacing commands (`\,` `\;` `\:` `\!` `\ `) are now collected as a single
`space` token before any operand scanning runs.

This is separate from `keep_spacing()` above: that one is about what survives
the trip through MathML and the XSLT, this one is about tokenising the LaTeX
source before the fraction builder ever sees it.

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

## Checking a formula: what each renderer can and cannot tell you

Three different questions get asked about a converted formula, and only one of
them needs Word. Getting the split wrong wastes hours in both directions —
staring at LibreOffice output for a defect it cannot show, or shipping on its
word for a defect only Word reveals.

| Question | Tool | Trustworthy? |
|---|---|---|
| Did an equation object get created at all? | `docx_read.py --runs`, OMML count | yes — it reads the XML |
| Is the structure legal? | `docx_validate.py`, `assert_legal()` | yes |
| Does the equation *look* right? | `docx_preview.py` (LibreOffice) | **no** — see below |
| Does the delivered page look right? | `pdf_export.ps1` (Word COM) | yes, at 150 dpi and above |

### LibreOffice cannot see `m:sty` at all

The upright pass above writes `m:rPr/m:sty val="p"`. LibreOffice **ignores that
element**, so an italic `max` and an upright `max` rasterise to the *same pixels*
through `docx_preview.py`: measured on one document, the XOR count between the two
regions was **0** at every DPI tried, while the same pair through `pdf_export.ps1`
differed in several hundred to several thousand pixels depending on the
rasterisation DPI (371 at 72 dpi, 1195 at 150 dpi on a 1275x1650 page). A preview
screenshot therefore cannot confirm or refute the upright fix — it renders both
forms with the same slant, so a *correct* document looks broken and a *broken* one
looks fine.

The same sample, rasterised through `pdf_export.ps1` at five DPIs:

| Rasterisation DPI | 72 | 96 | 150 | 200 | 300 |
|---|---|---|---|---|---|
| `pdf_export.ps1` (Word COM), differing pixels | 371 | 556 | 1195 | 1957 | 3896 |
| `docx_preview.py` (LibreOffice), differing pixels | 0 | 0 | 0 | 0 | 0 |

**A pixel count is a property of the rasteriser, not of the document.** The same
pair of forms re-measured through a different pipeline — or the same pipeline at a
different DPI — gives a different number, so any figure quoted here has to carry
its page size and DPI or it will not reproduce. The figure that *is* stable is the
zero: LibreOffice returns 0 no matter how it is rasterised, because it never reads
the element in the first place.

### The sizes disagree too, and not symmetrically

Measured on one 1275x1650 page rendered at 150 dpi through both paths:

| Feature | LibreOffice | Word |
|---|---|---|
| `max` line height | 26 px | 21 px |
| integral sign, stroke height | 34 px | 35 px |
| integral sign, whole object incl. limits | 64 px | 58 px |
| the upper limit digit `1` | 10 px tall | 27 px band |
| cases closing brace, width x height | 7 x 70 px | 24 x 46 px |
| `pmatrix` alignment | centred | centred |

Read that table as *these are different renderers*, not *one of them is wrong*:
LibreOffice draws the integral smaller with its limits tucked closer, and draws the
cases brace as a thin tall stroke where Word draws a wide heavy one. Both are
internally consistent; only Word's is what the reader receives.

### The division of labour

1. **While iterating** — `docx_preview.py --pages DIR --dpi 150` for a fast look.
   Use it to confirm the formula exists, sits in the right paragraph, and has not
   been promoted to a display equation. Do not judge glyph shape, upright letters,
   integral size, or brace height from it.
2. **At sign-off** — export the same `.docx` through Word COM and read that PDF:

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File scripts\pdf_export.ps1 `
       -DocxPath out.docx -PdfPath out.pdf
   ```

   `pdf_export.ps1` takes `-DocxPath` (required) and `-PdfPath` (omit it and the PDF
   lands beside the input under the same stem), opens the document read-only through
   `Documents.Open`, and calls `ExportAsFixedFormat($pdf, 17)`. `tools/word_accept.py`
   wraps the same path when you want the result as JSON — `--pdf`, `--expect-pages`,
   `--expect-tables`, and exactly one JSON object on stdout.
3. **On the PDF** — zoom into the formula region. Anything about glyph shape,
   upright letters, operator size, or delimiter height is decided here and nowhere
   else.

> **A scanned textbook's printed page number is not the PDF page index.** The front
> matter (cover, preface, contents) usually carries its own numbering, so a book
> page 137 can be PDF page 149 or 131. Locate the formula by its equation number or
> the section heading around it, never by adding a fixed offset to the printed
> number — the offset changes at every section break, and a formula checked against
> the wrong page looks wrong in a way no amount of re-rendering fixes.

## LaTeX in Python source must be a raw string

Every LaTeX string you write in Python is a **raw string**:

```python
tex = r"\frac{1}{2}"          # ✅ one backslash reaches LaTeX
tex = "\frac{1}{2}"           # ❌ \f is a form feed: U+000C + "rac{1}{2}"
```

This is not style. Python resolves `\` + letter inside an ordinary string literal
into a *control character* before the converter ever sees it, and the result
usually does not raise — it converts into a quietly wrong formula.

### Which commands are hit

Measured by evaluating each command name in a plain literal and inspecting the
resulting code point (`scripts`-side probe, all 79 names in the chain's own
vocabulary):

| Escape | Code point | Commands whose backslash disappears |
|---|---|---|
| `\a` | U+0007 BEL | `\alpha`, `\angle`, `\approx`, `\array` |
| `\b` | U+0008 BS | `\bar`, `\begin`, `\beta`, `\binom`, `\bmod`, `\bot`, `\boxed` |
| `\f` | U+000C FF | `\frac`, `\forall` |
| `\n` | U+000A LF | `\nabla`, `\neq`, `\notin`, `\nu` |
| `\r` | U+000D CR | `\ref`, `\rho`, `\right` |
| `\t` | U+0009 TAB | `\tag`, `\tan`, `\text`, `\tfrac`, `\theta`, `\tilde`, `\times`, `\top`, `\triangle` |
| `\v` | U+000B VT | `\varphi`, `\vec`, `\vee` |
| `\x` | — | `\xi` and friends: **`SyntaxError: (unicode error) 'unicodeescape' codec can't decode`** — the only failure loud enough to stop you |

32 names break this way and 4 more raise at parse time. Forty-three others —
`\sqrt`, `\int`, `\sum`, `\lim`, `\partial`, `\infty`, `\left`, `\mathbb`, `\mathrm`,
`\lambda`, `\sigma`, `\cdot` and their neighbours — survive by accident: the letter
after the backslash starts no valid escape, so Python passes the pair through
untouched. **Do not learn the table. Write `r"..."` and the question never
arises.**

### What actually happens downstream

Feeding the mangled strings to `latex_to_omml_lxml` (seven probes, one per escape
class):

| Source text reaching the converter | Result |
|---|---|
| U+000C + `rac{1}{2}` | converts, `m:t` = `rac12` — silent |
| U+0009 + `imes 3` | converts, `m:t` = `2imes3` — silent |
| U+000A + `eq b` | converts, `m:t` = `aeqb` — silent |
| U+000D + `ho V` | converts, `m:t` = `hoV` — silent |
| U+000B + `ec{F}` | converts, `m:t` = `ecF` — silent |
| U+0008 + `eta` | raises `XMLSyntaxError: PCDATA invalid Char value 8` |
| U+0007 + `lpha` | raises `XMLSyntaxError: PCDATA invalid Char value 7` |

Five of seven fail **silently** with a plausible-looking but wrong formula. That
asymmetry is the whole reason this is a hard rule rather than a preference: the
loud cases get fixed immediately, the quiet ones ship.

### The same rule in JSON and in a spec

A JSON string has its own escape layer, so the backslash is doubled there:

| Context | Written as | Meaning |
|---|---|---|
| Python raw string | `r"\frac{1}{2}"` | one backslash |
| JSON | `"\\frac{1}{2}"` | one backslash after JSON decoding |
| JSON, single backslash | `"\frac{1}{2}"` | U+000C + `rac{1}{2}` — the Python bug, in JSON |

`"\beta"` gives U+0008 + `eta`; `"\vee"`, `"\alpha"` and `"\xi"` make `json.loads`
raise `Invalid \escape`. A `docx_create` spec is JSON, so its `latex` blocks follow
the JSON column — that is why every example in this manual is spelled
`"latex": "\\frac{a}{b} = c"`.

The chain's own `--tex` output and `--report` tables are written with the literal
backslash, so reading them back into a Python literal still needs `r"..."`.

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
