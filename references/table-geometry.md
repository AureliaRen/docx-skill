# Table geometry: why your column widths are ignored

## The three places a width can live

OOXML gives you three separate, overlapping places to express how wide a column
should be:

| Element | Where | Meaning |
|---|---|---|
| `<w:tblW>` | `w:tblPr` | overall table width |
| `<w:gridCol w:w="..."/>` | `w:tblGrid` | the grid's column definitions |
| `<w:tcW w:w="..."/>` | each `w:tc` | that cell's preferred width |

All widths are `dxa` — twentieths of a point. `1 cm = 566.93 dxa`.

## What the renderer actually does

Measured against Word and WPS, 2026-10-27:

**`tblLayout` absent (autofit).** Each cell's `<w:tcW>` acts as a *hint*.
Columns come out proportional to the `tcW` values. Setting `tcW` therefore
"works", and the grid is largely ignored. This is the mode people accidentally
rely on.

**`tblLayout w:type="fixed"`.** The renderer lays out **purely from
`<w:tblGrid><w:gridCol>` and ignores `tcW` entirely**. If the grid still holds
the values Word generated on creation — often all equal — every column comes out
the same width no matter what you set in `tcW`.

That is the trap. Fixing `tblLayout` to get predictable output *immediately*
breaks `tcW`, because the grid is now the only thing being read.

## The rule

Keep all three consistent. Whichever element is authoritative, the other two must
agree with it, or you get behaviour that changes depending on which toggle
happened last.

## Using `table_width.py`

```bash
# What is the current state?
python scripts/table_width.py report unpacked/

# Copy existing tcW values into the grid (cheap fix, keeps current ratios)
python scripts/table_width.py sync unpacked/

# Pin exact widths in centimetres (sets tblW + gridCol + tcW together)
python scripts/table_width.py set unpacked/ --widths 2.4,4.2,4.2
```

`report` prints the grid, the `tcW` values, and whether they disagree — start
there, because the right fix depends on which mismatch you actually have.

The script touches horizontal geometry only. Row heights are left alone.

## Common symptoms

| Symptom | Cause | Fix |
|---|---|---|
| every column equal | `fixed` layout + untouched grid | `set --widths` |
| widths shift when you edit a cell | autofit recomputing from `tcW` | set `fixed` then `sync` |
| first column correct, rest off | grid and `tcW` disagree | `report`, then `sync` |
| total width exceeds page | `tblW` inconsistent with grid sum | `set` (recomputes `tblW`) |
