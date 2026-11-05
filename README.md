# docx

Create, read, edit, validate, and rewrite Word documents from the command line.

A skill for agents and humans alike: small, single-purpose Python CLIs that each
do one thing to a `.docx` and print JSON. No Office installation required for the
core operations — only `python-docx` and `lxml`.

## Why another docx toolkit

Most document automation either hides the XML completely (and then can't express
what you need) or drops you into raw OOXML with no guard rails. This one sits in
between: each script is a narrow, testable operation with an argparse surface
you can call from a shell, a Makefile, or an agent.

Two things it does that are hard to find elsewhere:

- **Native Office Math from LaTeX.** Converts equations into real OMML —
  editable in Word, not an image and not a text approximation. Works in bulk
  across a whole document, or one equation at a time.
- **Table widths that actually stick.** Explains and fixes the
  `tblLayout` / `gridCol` / `tcW` interaction that makes column widths appear to
  be ignored.

## Install

```bash
git clone https://github.com/AureliaRen/docx-skill.git
cd docx-skill
python -m pip install -r requirements.txt
```

Python 3.10 or newer.

Optional extras:

| Feature | Extra requirement |
|---|---|
| Formula conversion | `latex2mathml` + a local Microsoft Office (for `MML2OMML.XSL`) |
| Template apply (creating a doc from a template) | Node.js with the `docx` package |

Legacy `.doc` (binary) files are not supported — save as `.docx` first, or use
LibreOffice externally.

## Use

Every script is a CLI. Start with `--help`:

```bash
python scripts/docx_create.py --help
python scripts/docx_read.py --help
```

A minimal round trip:

```bash
# create
cat > spec.json <<'JSON'
{"blocks": [
  {"type": "heading",   "text": "Report", "level": 1},
  {"type": "paragraph", "text": "Ohm's law: U=IR"}
]}
JSON
python scripts/docx_create.py spec.json report.docx

# inspect
python scripts/docx_read.py report.docx --structure

# convert the formula into native Office Math
python scripts/docx_mathify.py report.docx -o report-math.docx --report conv.md
```

Full command table in [SKILL.md](SKILL.md).

## Layout

```
scripts/     flat CLI modules, imported by bare name
references/  longer notes loaded on demand
tests/       pytest suite
examples/    sample specs
```

All scripts share one import discipline: each puts its own directory on
`sys.path` before importing a sibling. That means `python scripts/x.py`,
`python x.py` from inside `scripts/`, and an absolute path from a different
directory all work — no `__init__.py` chain to get wrong.

## Testing

```bash
python -m pytest tests/ -v
```

## License

MIT — see [LICENSE](LICENSE).

The formula engine reuses Microsoft's `MML2OMML.XSL` from a local Office
installation at runtime; that file is not redistributed here.

## Acknowledgments

The base CRUD, comment, revision, and template CLIs are adapted from the
clean-room MIT implementations in
[NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent)
(`skills/productivity/microsoft-office`), originally written from functional
specs against `python-docx` and friends.
