# Revisions and comments

## Two different mechanisms

Word has two independent annotation systems, and they are easy to confuse:

| | Tracked changes | Comments |
|---|---|---|
| Stored in | `word/document.xml` | `word/comments.xml` |
| Elements | `<w:ins>`, `<w:del>` | `<w:comment>` + `<w:commentRangeStart/End>` |
| Semantics | the document text *contains* both versions | the text is unchanged; a note hangs off it |
| Accepting | choose which version stays | delete the anchor and the note |

A document can have both at once, and they interact: a comment can be anchored
inside a deleted range, in which case accepting the deletion orphans the comment.

## Revisions

`docx_revisions.py` treats revisions as a list you can act on:

```bash
python scripts/docx_revisions.py list in.docx          # JSON, with w:id per change
python scripts/docx_revisions.py accept-all in.docx
python scripts/docx_revisions.py reject-all in.docx
python scripts/docx_revisions.py accept in.docx --id 3 # one change
python scripts/docx_revisions.py reject in.docx --id 7
```

Each command takes `-o/--output`; without it the file is edited in place.

**Revision ids are position-dependent.** They are assigned by the editing client
and are not stable across a round trip through Word. If you plan to accept
specific changes, list first, act immediately, and do not reuse ids from an older
run.

## Comments

```bash
python scripts/docx_comments.py list in.docx
python scripts/docx_comments.py add in.docx --target "the phrase to anchor to" --text "note"
python scripts/docx_comments.py add in.docx --target "..." --text "..." --author "Reviewer" --initials RV
python scripts/docx_comments.py delete in.docx --id 1
```

Anchoring works by locating the target text in the document and wrapping it in
`<w:commentRangeStart>` / `<w:commentRangeEnd>` plus a `<w:commentReference>` run.
Consequences worth knowing:

- The target must appear exactly once. Ambiguous text fails rather than guessing
  which occurrence you meant.
- If the text is split across several runs (common when formatting changes
  mid-phrase), the matcher has to span runs. That is why editing a document after
  adding a comment, then editing it again, can shift anchors.
- Adding a comment rewrites the whole document part, so do comment work before
  heavy reformatting.

## Checking whether a document is clean

`docx_validate.py` reports issues as JSON. A document with no revisions and no
comments validates with `{"ok": true, "issues": []}`.

Before delivering a document that went through review, confirm no revisions
remain — an unaccepted tracked change is visible to the recipient as markup.

## Interaction with the other scripts

- `docx_edit.py` edits the *current* text, which for a region with tracked
  changes includes both the original and the revision. Accept or reject first,
  then edit.
- `docx_mathify.py` converts runs in place and preserves revision markers, but a
  formula split across an insert/delete boundary will not convert — the span
  detector sees two fragments. Resolve revisions first.
