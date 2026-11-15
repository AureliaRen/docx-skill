# Editing traps: six things that fail silently

Every trap here has the same shape: the script exits `0`, stdout says `{"ok":
true}`, stderr is empty — and the document that comes out is wrong in a way only
a human opening it in Word will notice. That is what makes them expensive.

The rule that follows from all six: **`exit 0` is not evidence that the edit
did what you meant it to do.** Read the output document back before you hand it
over.

Two of the six are fixed in the scripts (a warning or an error now fires). The
other four are not fixable in a library — they are mistakes in how a caller uses
it, so the fix lives here.

---

## 1. `cell.text = "..."` flattens a multi-run cell

### What happens

`python-docx`'s `_Cell.text` setter is not a text replacement. Its docstring
says so plainly:

> Any existing content or revisions are replaced.

Its implementation is `clear_content()` → `add_p()` → `add_r()` → `r.text = text`.
Everything that made the cell readable goes: paragraph alignment, run font, run
size, bold/italic, and any `vertAlign` sub/superscript. A three-run cell becomes
one run; a two-paragraph cell becomes one paragraph.

Reproduced on a 1×1 table, 2 paragraphs, 3 runs (second run `subscript`), first
paragraph `CENTER`, Times New Roman 14 pt, after writing `"d0=5"`:

| | before | after |
|---|---|---|
| paragraphs in cell | 2 | **1** |
| runs | 3 | **1** |
| `vertAlign` on run 2 | `subscript` | **gone with the run** |
| paragraph alignment | `CENTER` | **`None`** |
| run size / font | 14 pt / Times New Roman | **`None` / `None`** |
| exit code / stderr / stdout | — | **`0` / empty / `{"ok": true}`** |

### What the script does now

`docx_edit.py set-cell` no longer calls that setter. Its default is
`mode: "preserve"`: it writes into the first *writable* run and leaves the
paragraph alone. `--replace-cell` restores the old clearing behaviour for the
cases where you actually want a clean cell — use it deliberately, not by
accident.

The multi-run case still loses the extra runs (a flat string cannot express
them), but it now tells you: a `warning:` line on stderr, a `warnings` array in
the JSON, and `mode: "preserve"` rather than a bare `ok`.

### What you have to do instead

Read the cell first:

```bash
python scripts/docx_read.py in.docx --runs
```

`mixed_format: true` on a paragraph is the signal that a flat string cannot
carry this cell. When you need the formatting back, edit through `python-docx`
directly:

```python
import copy
from docx import Document

doc = Document("in.docx")
cell = doc.tables[0].cell(row=1, col=2)

# Keep the shape: take the run that already looks right, clone it, refill it.
template_run = cell.paragraphs[0].runs[0]
new_run = copy.deepcopy(template_run._r)
new_run.find(qn("w:t")).text = "d0=5"          # keeps rPr, incl. vertAlign
cell.paragraphs[0]._p.append(new_run)
cell.paragraphs[0]._p.remove(template_run._r)

doc.save("out.docx")
```

`copy.deepcopy` on the `w:r` **element** (not on the `Run` wrapper) is the point:
it carries `w:rPr` with it, so font, size, bold and `vertAlign` survive the
replacement. Build a run from scratch and you are back to guessing at formatting.

If the cell needs several differently formatted runs, clone the template run
once per run you need and refill each clone. Do not append plain strings and
then try to fix the formatting afterwards — `w:rPr` ordering rules ([see
`font-scheme.md`](font-scheme.md)) make retrofitting fragile.

---

## 2. A stray `vertAlign` turns a whole line into a subscript

### What happens

`vertAlign` lives on `w:rPr` of a *run*, but it renders as if it applied to the
line — a run carrying `subscript` makes its text sit below the baseline, and if
that run holds the start of a line, the line reads as one long subscript. This
is common in templates built from a subscripted formula whose author then kept
typing in the same run.

### What the script does now

`docx_edit.py` has two helpers:

- `_writable_run(run)` — tells you whether a run may be overwritten at all: it
  returns `False` for field runs (`w:fldChar`, `w:instrText`), drawings, and
  pictures.
- `_strip_vert_align(run)` — removes `w:rPr/w:vertAlign` and returns the value it
  removed, so the caller can report it.

In `preserve` mode, when the receiving run carries `vertAlign`, it is stripped
before the new text is written — otherwise the replacement text inherits the old
subscript. That is a deliberate change of the run's formatting, so it is reported
as a warning rather than done silently.

### What you have to do

When you want the sub/superscript kept (a variable name like `v₀`), do not
write into the subscripted run — write into the plain one and leave the
subscripted run alone. `docx_read.py --runs` shows you which run is which.

---

## 3. `inline_shapes` is a live collection — index-based access shifts

### What happens

`document.inline_shapes` is a *view* over the live XML, not a snapshot. Delete or
insert an inline shape and every later index moves. Code that captures
`shapes[i]` in one loop and uses `i` in a later step (or after any edit) silently
addresses the wrong image — and if the document mixes a school logo with content
figures, "the wrong image" means you just edited the logo.

### What you have to do

Match by identity, not by position. The shape's relationship id is stable across
edits:

```python
from docx.oxml.ns import qn

def embedded_images(paragraph):
    """Yield (rId, blob) for every inline image in a paragraph."""
    for blip in paragraph._p.iter(qn("a:blip")):
        rid = blip.get(qn("r:embed"))
        if rid:
            yield rid, paragraph.part.related_parts[rid].blob
```

Resolve the part you care about from the `rId` of the shape you inspected, and
re-resolve it after every structural edit. If you must iterate, collect the
`rId` list once, then re-look-up by `rId` on each use.

Same rule for `part.related_parts`: the dict is stable per `rId`, so pin to the
`rId` and never to insertion order.

---

## 4. An empty `<w:drawing/>` makes Word refuse to open the file

### What happens

Removing an image by deleting its `a:blip` — or by clearing the wrong node while
trying to remove a shape — can leave behind an empty drawing container:

```xml
<w:r><w:drawing/></w:r><w:r><w:t xml:space="preserve"> after</w:t></w:r>
```

That is not a valid drawing: `w:drawing` must contain exactly one `wp:inline` or
`wp:anchor`. Word does not repair it — `Documents.Open()` raises:

```
com_error (-2147352567, '发生意外。', ..., 'Word 在试图打开文件时遇到错误。...',
           'wdmain11.chm', 24601, ...)
```

Until P13, `docx_validate.py` reported `{"ok": true, "issues": []}` for such a
file and exited `0`, because its checks only covered package structure, rels,
referenced ids, media magic bytes and style existence — an empty drawing
references nothing, so nothing looked at it.

### What the script does now

`docx_validate.py` checks every `w:drawing` for a `wp:inline` or `wp:anchor`
child and reports:

```json
{"ok": false, "issues": [{"severity": "error", "code": "empty-drawing",
  "detail": "document.xml: <w:drawing> has neither wp:inline nor wp:anchor; Word refuses the file"}]}
```

with exit code `1`. The gate is "does Word accept it", not "does it look
plausible": the check was verified against five real documents
(`in.docx`, `out.docx`, `out_tnr.docx`, `final_test.docx`, `fixture.docx`) and
all five still come back `{"ok": true, "issues": []}`.

The same pass also flags a block element sitting directly inside a run
(`<w:p>` or `<w:tbl>` inside `<w:r>`, code `invalid-nesting`). That shape comes
out of hand-editing XML or from copying nodes across the block/inline boundary;
python-docx opens it fine, and Word either repairs it silently or refuses the
file. Only those two tags are flagged — a wider net starts reporting legal
`mc:AlternateContent` fallbacks, which are paragraphs inside something that
looks like a run but is not one.

### What you have to do

When removing an image, remove the whole enclosing `w:r` (or the whole
`w:drawing`), not just the `a:blip` inside it. Then run the validator before
declaring the edit finished:

```bash
python scripts/docx_validate.py out.docx
```

A `0` exit is only meaningful now that this check exists — before P13 it was
not evidence of anything about drawings.

---

## 5. A file open in WPS or Word blocks the write with `PermissionError`

### What happens

Word and WPS both hold the `.docx` open with a share mode that denies writes
while the document is being edited. `Document.save()` against an open file
raises:

```
PermissionError: [Errno 13] Permission denied: 'report.docx'
```

The message names the file, not the cause, so it reads like a path problem.

### What you have to do

- Check the obvious first: `Get-Process WINWORD, wps -ErrorAction SilentlyContinue`,
  and close the document in the editor (not just the window — WPS keeps a
  background process holding recent files).
- Write to a *different* path (`-o out.docx`) while iterating. The skill's CLIs
  default `-o` to overwriting the input, which is exactly the case that collides
  with an open editor. Pass `-o` explicitly when a human may have the file open.
- A stale `~$report.docx` lock file next to the document means an editor crashed;
  the lock is not enforced by the OS but Word may still prompt. Deleting a `~$`
  file is safe when no editor is running.
- Retry with a short backoff rather than immediately: an editor finishing a
  save releases the handle within a second or two.

---

## 6. Console output is mojibake on a Chinese Windows console (cp936)

### What happens

Chinese Windows consoles default to cp936. A script that prints a character cp936
cannot encode — an em dash, `→`, `≥`, a typographic quote, a rare CJK glyph —
raises:

```
UnicodeEncodeError: 'gbk' codec can't encode character '\u2192' in position ...
```

and if the exception fires mid-write the document may already be saved while the
report is lost, so you cannot tell what happened. On the read side the mirror
image occurs: a UTF-8 file read as cp936 produces `锟斤拷` or `口口`.

### What you have to do

Set the interpreter's I/O encoding, do not set the console codepage:

```powershell
$env:PYTHONIOENCODING = 'utf-8'
python -X utf8 scripts/docx_read.py in.docx --runs
```

`-X utf8` also switches the filesystem encoding, which matters when the path
itself contains CJK characters. Both are per-process, so they cannot break
anything else on the machine.

When writing a report file, be explicit rather than relying on the locale:

```python
Path("report.txt").write_text(text, encoding="utf-8")
```

If output is already garbled, do not guess at re-encoding — load the
`garbled-text-checker` skill, which identifies the specific corruption type
first. Text containing U+FFFD has lost information and cannot be recovered; go
back to the source.

---

## Checklist before you hand a document over

1. `python scripts/docx_read.py out.docx --runs` — did the runs and their
   formatting survive where you meant them to?
2. `python scripts/docx_validate.py out.docx` — exit `0`, and specifically no
   `empty-drawing` issue.
3. Open it once in Word or WPS. A file that passes both checks can still look
   wrong (column widths, page breaks), and only the renderer can tell you.
4. For a template fill: diff the run inventory before and after, not just the
   visible text. Text equality does not imply formatting equality.
