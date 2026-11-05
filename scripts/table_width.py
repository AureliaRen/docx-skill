"""Set fixed table geometry on an unpacked .docx: tblW + gridCol + tcW together.

How OOXML actually resolves column widths (measured with Word/WPS, 2026-10-27):

  * ``tblLayout`` absent (autofit) → renderer honours each cell's ``<w:tcW>``
    as a width *hint*; columns come out proportional to tcW.
  * ``tblLayout w:type="fixed"``  → renderer lays out purely from
    ``<w:tblGrid><w:gridCol>`` and **ignores tcW**.  If the grid is the default
    equal split, every column renders the same width no matter what tcW says.

That second case is the trap: python-docx's ``table.autofit = False`` writes
``tblLayout=fixed``, so setting ``cell.width`` alone silently produces equal
columns.  Setting all three places keeps them consistent:

    <w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="TOTAL"/>
              <w:tblLayout w:type="fixed"/>
           <w:tblGrid><w:gridCol w:w="W1"/><w:gridCol w:w="W2"/>...
           <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="W1"/>

Usage:
    # report the current geometry of every table
    python table_width.py report unpacked/

    # set column widths in cm, uniformly
    python table_width.py set unpacked/ --table 1 --widths 2.4,4.2,2.4,5.6

    # same width for all columns of all tables
    python table_width.py set unpacked/ --all --widths 3.5,3.5,3.5

    # copy the widths already set on cells into the grid, then pin everything
    python table_width.py sync unpacked/

    # split the text width evenly across all columns of every table
    python table_width.py auto unpacked/ --content-width 14.6

Notes:
  * ``--content-width`` is the usable text width in cm (A4 minus margins);
    ``auto`` splits it evenly across columns.
  * All w:w values here are dxa (twentieths of a point); 1 cm = 566.93 dxa.
  * Row heights are left alone; this script only touches horizontal geometry.
"""
import os as _os, sys as _sys; _sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix
import argparse
import sys
from pathlib import Path

from docx_utils import (
    ET, wtag, register_all, parse, write, fix_root,
    insert_ordered, cm_to_dxa, TBLPR_ORDER,
)


def _tables(root):
    return list(root.iter(wtag("tbl")))


def _ensure_tblpr(tbl):
    tblPr = tbl.find(wtag("tblPr"))
    if tblPr is None:
        tblPr = ET.Element(wtag("tblPr"))
        tbl.insert(0, tblPr)
    return tblPr


def _ensure_grid(tbl, n_cols):
    grid = tbl.find(wtag("tblGrid"))
    if grid is None:
        grid = ET.Element(wtag("tblGrid"))
        tblPr = tbl.find(wtag("tblPr"))
        idx = list(tbl).index(tblPr) + 1 if tblPr is not None else 0
        tbl.insert(idx, grid)
    cols = grid.findall(wtag("gridCol"))
    while len(cols) < n_cols:
        gc = ET.Element(wtag("gridCol"))
        grid.append(gc)
        cols = grid.findall(wtag("gridCol"))
    while len(cols) > n_cols:
        grid.remove(cols[-1])
        cols = grid.findall(wtag("gridCol"))
    return grid


def _row_cells(tr):
    return tr.findall(wtag("tc"))


def _cell_span(tc):
    """Horizontal span of a cell (gridSpan), default 1."""
    tcPr = tc.find(wtag("tcPr"))
    if tcPr is None:
        return 1
    gs = tcPr.find(wtag("gridSpan"))
    if gs is None:
        return 1
    try:
        return max(1, int(gs.get(wtag("val"), "1")))
    except ValueError:
        return 1


def _col_count(tbl):
    """Column count from the widest row, honouring gridSpan."""
    n = 0
    for tr in tbl.findall(wtag("tr")):
        spans = sum(_cell_span(tc) for tc in _row_cells(tr))
        n = max(n, spans)
    grid = tbl.find(wtag("tblGrid"))
    if grid is not None:
        n = max(n, len(grid.findall(wtag("gridCol"))))
    return n


def _set_tbl_width(tblPr, dxa):
    el = tblPr.find(wtag("tblW"))
    if el is None:
        el = ET.Element(wtag("tblW"))
        insert_ordered(tblPr, el, TBLPR_ORDER)
    el.set(wtag("type"), "dxa")
    el.set(wtag("w"), str(dxa))


def _set_fixed_layout(tblPr):
    el = tblPr.find(wtag("tblLayout"))
    if el is None:
        el = ET.Element(wtag("tblLayout"))
        insert_ordered(tblPr, el, TBLPR_ORDER)
    el.set(wtag("type"), "fixed")


def _apply_widths(tbl, widths_dxa):
    """Write gridCol + every tcW from an explicit per-column list."""
    n = len(widths_dxa)
    grid = _ensure_grid(tbl, n)
    for gc, w in zip(grid.findall(wtag("gridCol")), widths_dxa):
        gc.set(wtag("w"), str(w))

    tblPr = _ensure_tblpr(tbl)
    _set_tbl_width(tblPr, sum(widths_dxa))
    _set_fixed_layout(tblPr)

    for tr in tbl.findall(wtag("tr")):
        ci = 0
        for tc in _row_cells(tr):
            span = _cell_span(tc)
            w = sum(widths_dxa[ci:ci + span]) if ci < n else widths_dxa[-1]
            tcPr = tc.find(wtag("tcPr"))
            if tcPr is None:
                tcPr = ET.Element(wtag("tcPr"))
                tc.insert(0, tcPr)
            tcW = tcPr.find(wtag("tcW"))
            if tcW is None:
                tcW = ET.Element(wtag("tcW"))
                insert_ordered(tcPr, tcW, "tcPr")
            tcW.set(wtag("type"), "dxa")
            tcW.set(wtag("w"), str(w))
            ci += span


def _read_tcw(tbl):
    """Per-column widths from the first row's <w:tcW> (dxa or None).

    This is what a python-docx author actually set with ``cell.width`` — the
    value worth preserving when repairing a table.
    """
    for tr in tbl.findall(wtag("tr")):
        vals = []
        for tc in _row_cells(tr):
            span = _cell_span(tc)
            tcPr = tc.find(wtag("tcPr"))
            tcW = tcPr.find(wtag("tcW")) if tcPr is not None else None
            v = None
            if tcW is not None and tcW.get(wtag("type")) in (None, "dxa"):
                try:
                    v = int(tcW.get(wtag("w")))
                except (TypeError, ValueError):
                    v = None
            vals.append(v if span == 1 else (v / span if v else None))
        if any(vals):
            return vals
    return []


def _read_grid(tbl):
    """Per-column widths from <w:tblGrid> (dxa or None)."""
    grid = tbl.find(wtag("tblGrid"))
    if grid is None:
        return []
    vals = []
    for gc in grid.findall(wtag("gridCol")):
        try:
            vals.append(int(gc.get(wtag("w"))))
        except (TypeError, ValueError):
            vals.append(None)
    return vals


def _read_widths(tbl):
    """Best-effort per-column widths already present in the file (dxa or None)."""
    vals = _read_grid(tbl)
    if vals and any(vals):
        return vals
    return _read_tcw(tbl)


def _report(tbl, idx):
    rows = tbl.findall(wtag("tr"))
    n = _col_count(tbl)
    tblPr = tbl.find(wtag("tblPr"))
    layout = tblPr.find(wtag("tblLayout")) if tblPr is not None else None
    tblW = tblPr.find(wtag("tblW")) if tblPr is not None else None
    print(f"table {idx}: {len(rows)} rows x {n} cols")
    print(f"  tblW      = {tblW.get(wtag('w')) if tblW is not None else '(none)'}"
          f"  type={tblW.get(wtag('type')) if tblW is not None else '-'}")
    print(f"  tblLayout = {layout.get(wtag('type')) if layout is not None else '(none)'}")
    ws = _read_widths(tbl)
    if ws:
        cm = [f"{(w / 1440 * 2.54):.2f}" if w else "?" for w in ws]
        print(f"  gridCol   = {ws}  ({', '.join(cm)} cm)")
        print(f"  total     = {sum(w for w in ws if w)} dxa"
              f"  ({(sum(w for w in ws if w) / 1440 * 2.54):.2f} cm)")
    else:
        print("  gridCol   = (none set) — widths will not be honoured")

    # tcW vs gridCol: only worth flagging when the gap is visible.
    # Rounding from cm→dxa is ±1; a real mismatch (e.g. equal-split grid vs
    # explicit widths) is off by hundreds.
    #
    # A cell with gridSpan=N covers N grid columns and its tcW is the sum of
    # theirs, so comparing it against a single grid column is meaningless --
    # that produces a false "mismatch" on any table with merged cells, and
    # following the advice would corrupt the grid. Compare only spanning-1
    # cells, and advance the column cursor by the span.
    TOL = 20   # dxa ≈ 0.35 mm — far below anything a reader would notice
    first = rows[0] if rows else None
    layout_fixed = (layout is not None and layout.get(wtag("type")) == "fixed")
    mismatch = False
    if first is not None:
        ci = 0
        for tc in _row_cells(first):
            span = _cell_span(tc)
            tcPr = tc.find(wtag("tcPr"))
            tcW = tcPr.find(wtag("tcW")) if tcPr is not None else None
            have = tcW.get(wtag("w")) if tcW is not None else None
            if span == 1:
                want = ws[ci] if ci < len(ws) else None
                if have and want and abs(int(have) - int(want)) > TOL:
                    mismatch = True
                    dh, dw = int(have), int(want)
                    print(f"  ! col {ci + 1}: tcW={dh} ({dh / 1440 * 2.54:.2f} cm)"
                          f" != gridCol={dw} ({dw / 1440 * 2.54:.2f} cm)")
            else:
                covered = ws[ci:ci + span]
                total = sum(w for w in covered if w)
                note = "" if have is None else f", tcW={int(have)}"
                print(f"  . col {ci + 1}-{ci + span}: merged cell"
                      f" (gridSpan={span}, grid total={total}{note})")
            ci += span
    if mismatch and layout_fixed:
        print("  ! tblLayout=fixed ignores tcW — columns render from gridCol only."
              "  Run 'sync' to copy tcW into the grid.")


def main():
    ap = argparse.ArgumentParser(description="Fix table column widths in an unpacked docx")
    ap.add_argument("action", choices=["report", "set", "sync", "auto"])
    ap.add_argument("unpacked_dir")
    ap.add_argument("--table", type=int, default=None,
                    help="1-based table index (default: all)")
    ap.add_argument("--all", action="store_true", help="apply to every table")
    ap.add_argument("--widths", default=None,
                    help="comma-separated column widths in cm, e.g. 2.4,4.2,2.4,5.6")
    ap.add_argument("--content-width", type=float, default=14.6,
                    help="usable text width in cm for 'auto' (default 14.6 = A4 - 3.17cm margins)")
    args = ap.parse_args()

    doc_xml = Path(args.unpacked_dir) / "word" / "document.xml"
    if not doc_xml.is_file():
        print(f"error: {doc_xml} not found", file=sys.stderr)
        return 1

    register_all()
    tree = parse(doc_xml)
    root = tree.getroot()

    tables = _tables(root)
    if not tables:
        print("no tables found")
        return 0

    if args.action == "report":
        for i, t in enumerate(tables, 1):
            _report(t, i)
            print()
        return 0

    targets = []
    if args.table:
        if not (1 <= args.table <= len(tables)):
            print(f"error: table {args.table} out of range (1..{len(tables)})", file=sys.stderr)
            return 1
        targets = [tables[args.table - 1]]
    elif args.all or not args.table:
        targets = tables

    explicit = None
    if args.widths:
        try:
            explicit = [cm_to_dxa(float(x)) for x in args.widths.split(",") if x.strip()]
        except ValueError:
            print("error: --widths must be comma-separated numbers in cm", file=sys.stderr)
            return 1

    for ti, tbl in enumerate(targets, 1):
        n = _col_count(tbl)
        if args.action == "set":
            if not explicit:
                print("error: 'set' requires --widths", file=sys.stderr)
                return 1
            widths = explicit[:n] + [explicit[-1]] * max(0, n - len(explicit))
        elif args.action == "sync":
            # prefer the widths the author set on cells; fall back to the grid
            widths = _read_tcw(tbl) or _read_grid(tbl)
            if not widths or any(w is None for w in widths):
                print(f"skip table {ti}: no usable widths to sync from", file=sys.stderr)
                continue
            widths = [int(w) for w in widths]
            widths = widths[:n] + [widths[-1]] * max(0, n - len(widths))
        else:  # auto
            if explicit:
                widths = explicit[:n] + [explicit[-1]] * max(0, n - len(explicit))
            else:
                each = cm_to_dxa(args.content_width / n)
                widths = [each] * n
        _apply_widths(tbl, widths)
        cm = ", ".join(f"{w / 1440 * 2.54:.2f}" for w in widths)
        print(f"table {ti}: set {n} cols -> {cm} cm (total "
              f"{sum(widths) / 1440 * 2.54:.2f} cm, fixed layout)")

    write(tree, doc_xml)
    fix_root(doc_xml)
    print(f"written: {doc_xml}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
