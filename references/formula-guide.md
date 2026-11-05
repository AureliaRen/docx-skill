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
```

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

## Merged scripts

Word does not store "sub 0, then letter L" as two scripts. It stores a run whose
`vertAlign` says subscript and whose text is `0L`. Converting that literally
gives `\omega_{0L}` — one subscript containing two tokens — when the author meant
`\omega_0 L`.

`repair_merged_scripts()` splits these apart heuristically: it recognises when
the tail of a script is actually a following token, and honours the cases where
it genuinely is one script (`U^{2}`, `10^{-6}`).

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
