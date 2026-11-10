#!/usr/bin/env python3
# MIT License. Shared helpers for the docx skill scripts.
"""Shared helpers: paragraph iteration and run-preserving text replacement."""
from __future__ import annotations

# --- repo bootstrap: make sibling imports work from any cwd ---------------
import _bootstrap  # noqa: F401 — UTF-8 stdio on Windows
# --------------------------------------------------------------------------


def iter_all_paragraphs(doc, include_headers_footers: bool = True):
    """Yield every paragraph in body, tables (recursively), headers, footers."""
    yield from _iter_container(doc)
    if include_headers_footers:
        for section in doc.sections:
            for part in (
                section.header, section.footer,
                section.first_page_header, section.first_page_footer,
                section.even_page_header, section.even_page_footer,
            ):
                if part is not None:
                    yield from _iter_container(part)


def _iter_container(container):
    yield from container.paragraphs
    for table in container.tables:
        yield from _iter_table(table)


def _iter_table(table):
    for row in table.rows:
        for cell in row.cells:
            yield from cell.paragraphs
            for nested in cell.tables:
                yield from _iter_table(nested)


def iter_part_roots(doc):
    """Yield the XML root of the body plus every header/footer part."""
    yield doc.element.body
    seen = set()
    for section in doc.sections:
        for part in (
            section.header, section.footer,
            section.first_page_header, section.first_page_footer,
            section.even_page_header, section.even_page_footer,
        ):
            if part is not None and id(part._element) not in seen:
                seen.add(id(part._element))
                yield part._element


def replace_in_paragraph(para, old: str, new: str) -> int:
    """Replace `old` with `new` in a paragraph, preserving run formatting.

    Strategy: first replace occurrences fully contained in a single run
    (formatting fully preserved). Then handle needles that span runs; those runs
    are collapsed and the replacement inherits the formatting of the run where
    the match starts.

    Returns the number of replacements made.

    The total is budgeted from the occurrence count in the *original* text. That
    matters when `new` contains `old` ("cat" -> "catcat", or a template token
    whose value is the token itself): rewriting creates fresh matches, so any
    loop driven by "does the needle still appear?" never terminates. Budgeting
    up front makes the operation finite and idempotent in intent.
    """
    if not old or old not in para.text:
        return 0

    original_count = para.text.count(old)
    count = 0

    # Pass 1: replacements fully inside a single run.
    for run in para.runs:
        if old in run.text:
            hits = run.text.count(old)
            run.text = run.text.replace(old, new)
            count += hits

    # Pass 2: occurrences split across run boundaries.
    budget = original_count - count
    while budget > 0:
        budget -= 1
        runs = para.runs
        full = "".join(r.text for r in runs)
        start = full.find(old)
        if start < 0:
            break
        end = start + len(old)
        # If this match sits wholly inside one run, Pass 1 would have taken it.
        # Skip ahead rather than re-collapsing the same text forever.
        pos = 0
        spans = []
        for i, r in enumerate(runs):
            r_start, r_end = pos, pos + len(r.text)
            if r_end > start and r_start < end:
                spans.append((i, max(start, r_start) - r_start,
                              min(end, r_end) - r_start))
            pos = r_end
        single = len(spans) == 1
        if single:
            # Should not happen after Pass 1, but guard against a rewrite loop.
            break
        first = True
        for i, cs, ce in spans:
            t = runs[i].text
            if first:
                runs[i].text = t[:cs] + new + t[ce:]
                first = False
            else:
                runs[i].text = t[:cs] + t[ce:]
        count += 1

    return count


def main(argv=None) -> int:
    """自述入口：本模块是库，被 import 用；单独运行只打印用法，退出码 0。"""
    import argparse

    ap = argparse.ArgumentParser(
        prog="docx_common",
        description="Shared helpers imported by the other docx scripts "
                    "(library module, nothing to execute at run time).",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="被 docx_edit.py / docx_template.py / docx_comments.py / "
               "docx_revisions.py import；单独运行只打印自述。",
    )
    ap.add_argument("--api", action="store_true",
                    help="打印导出的函数清单（含首行 docstring）后退出")
    a = ap.parse_args(argv)
    if a.api:
        for fn in (iter_all_paragraphs, iter_part_roots, replace_in_paragraph):
            first = (fn.__doc__ or "").strip().splitlines()[0] if fn.__doc__ else ""
            print(f"{fn.__name__}  # {first}")
        return 0
    ap.print_help()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
