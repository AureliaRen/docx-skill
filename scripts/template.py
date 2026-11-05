r"""Extract/apply .docx formatting templates.

Usage:
    python template.py extract template.docx -o tpl.json
    python template.py apply tpl.json -o new.docx [--fix-formulas]
    python template.py apply tpl.json --to existing.docx [--fix-formulas]

--fix-formulas  Auto-convert _/^ markers to subscript/superscript in the output.
"""
from __future__ import annotations
import os as _os, sys as _sys; _sys.path.insert(0, _os.path.dirname(_os.path.abspath(__file__)))
import encoding_fix

import argparse, json, re, shutil, subprocess, sys, tempfile
from pathlib import Path
from xml.etree import ElementTree as ET
from docx_utils import *
SKILL = Path(__file__).resolve().parents[2]

_RPR_MAP = {"rFonts":"font","b":"bold","bCs":"boldCs","i":"italic","iCs":"italicCs",
            "u":"underline","color":"color","sz":"size","szCs":"sizeCs",
            "vertAlign":"vertAlign","kern":"kern"}
_PPR_MAP = {"jc":"align","outlineLvl":"outlineLevel"}
_STYLE_ID = {"normal":"Normal","heading 1":"Heading1","heading 2":"Heading2",
    "heading 3":"Heading3","heading 4":"Heading4","heading 5":"Heading5",
    "heading 6":"Heading6","heading 7":"Heading7","heading 8":"Heading8",
    "heading 9":"Heading9","title":"Title","subtitle":"Subtitle"}


# ---- parse helpers ----
def _simplify(attrs: dict) -> dict | str:
    """Strip namespace prefixes from dict keys."""
    c = {k.rsplit("}",1)[-1]: v for k, v in attrs.items()}
    return c.get("val", "1") if list(c) == ["val"] else c


def _parse_rpr(rpr: ET.Element | None) -> dict | None:
    if rpr is None: return None
    r = {}
    for e in rpr:
        tag = e.tag.rsplit("}",1)[-1]
        r[_RPR_MAP.get(tag, tag)] = _simplify(dict(e.attrib))
    return r or None


def _parse_ppr(ppr: ET.Element | None) -> dict | None:
    if ppr is None: return None
    r = {}
    for e in ppr:
        tag = e.tag.rsplit("}",1)[-1]
        s = _simplify(dict(e.attrib))
        if tag == "spacing": r["spacing"] = s
        elif tag == "ind": r["indent"] = s
        elif tag == "numPr": pass
        elif tag in _PPR_MAP: r[_PPR_MAP[tag]] = s if isinstance(s, dict) else s
        else: r[tag] = s
    return r or None


# ---- subprocess wrappers ----
def _env(): e = __import__('os').environ.copy(); e.setdefault("PYTHONIOENCODING","utf-8"); return e


def _run(*args: str) -> None:
    subprocess.run([sys.executable] + list(args), check=True,
                   capture_output=True, encoding="utf-8", errors="replace", env=_env())


def _run_node(script: str, output: str) -> None:
    subprocess.run(["node", script, output], check=True,
                   capture_output=True, encoding="utf-8", errors="replace", env=_env())


# ---- extract ----
def _extract(docx: Path) -> dict:
    with tempfile.TemporaryDirectory() as td:
        up = Path(td) / "u"
        _run(str(SKILL / "scripts/office/unpack.py"), str(docx), str(up))
        tpl: dict = {"page":{}, "defaults":{}, "styles":{}}

        sx = up / "word" / "styles.xml"
        if sx.exists():
            regw(); tree = parse(sx)
            dd = tree.getroot().find(wtag("docDefaults"))
            if dd is not None:
                rp = dd.find(wtag("rPrDefault"))
                if rp is not None and (r := _parse_rpr(rp.find(wtag("rPr")))): tpl["defaults"]["run"] = r
                pp = dd.find(wtag("pPrDefault"))
                if pp is not None and (r := _parse_ppr(pp.find(wtag("pPr")))): tpl["defaults"]["paragraph"] = r
            for s in tree.getroot().iter(wtag("style")):
                nm = s.find(wtag("name"))
                if nm is None: continue
                name = nm.get(wtag("val"),"")
                entry: dict = {"type": s.get(wtag("type"),"paragraph"), "styleId": s.get(wtag("styleId"),"")}
                rp = _parse_rpr(s.find(wtag("rPr")))
                pp = _parse_ppr(s.find(wtag("pPr")))
                if rp: entry["run"] = rp
                if pp: entry["paragraph"] = pp
                if bo := s.get(wtag("basedOn"),""): entry["basedOn"] = bo
                tpl["styles"][name] = entry

        dx = up / "word" / "document.xml"
        if dx.exists():
            regw(); tree = parse(dx)
            sect = tree.getroot().find(f"{wtag('body')}/{wtag('sectPr')}")
            if sect is not None:
                pg = sect.find(wtag("pgSz"))
                if pg is not None:
                    tpl["page"]["width"] = int(pg.get(wtag("w"),"11906"))
                    tpl["page"]["height"] = int(pg.get(wtag("h"),"16838"))
                    if o := pg.get(wtag("orient"),""): tpl["page"]["orientation"] = o
                pm = sect.find(wtag("pgMar"))
                if pm is not None:
                    tpl["page"]["margins"] = {a: int(pm.get(wtag(a))) for a in
                        ("top","right","bottom","left","header","footer") if pm.get(wtag(a))}
    return tpl


# ---- apply: new document (docx-js) ----
def _gen_js(tpl: dict) -> str:
    p = tpl.get("page",{}); d = tpl.get("defaults",{}); dr = d.get("run",{})
    w, h = p.get("width",12240), p.get("height",15840)
    mg = p.get("margins",{})

    df = {"ascii":"Times New Roman","hAnsi":"Times New Roman","eastAsia":"微软雅黑","cs":"Times New Roman"}
    if "font" in dr:
        f = dr["font"]
        df = {k: f.get(k, df[k]) for k in df} if isinstance(f, dict) else f
    ds = int(dr.get("size",{}).get("val",24)) if isinstance(dr.get("size"), dict) else int(dr.get("size", 24))

    font_js = "{" + ", ".join(f'{k}: "{v}"' for k, v in (df.items() if isinstance(df, dict) else {})) + "}" if isinstance(df, dict) else f'"{df}"'

    styles_js = []
    for name, s in tpl.get("styles",{}).items():
        if s.get("type") != "paragraph": continue
        sid = _STYLE_ID.get(name.lower(), name)
        ro = _run_opts(s.get("run",{}), df, ds)
        po = _para_opts(s.get("paragraph",{}))
        styles_js.append(f'    {{ id: "{sid}", name: "{name}", basedOn: "Normal", next: "Normal", quickFormat: true, run: {{ {", ".join(ro)} }}, paragraph: {{ {", ".join(po)} }} }}')

    # Pick heading styles for placeholder paragraphs (sorted by outline level)
    _NON_HEADING = {"normal", "normal (web)", "body text", "body text indent",
                    "plain text", "default paragraph font", "list paragraph"}
    heading_styles = []
    for name, s in tpl.get("styles", {}).items():
        if s.get("type") != "paragraph":
            continue
        if name.lower().strip() in _NON_HEADING:
            continue
        sid = s.get("styleId", "")
        level = _heuristic_outline_level(name)
        para_info = s.get("paragraph", {})
        ol = para_info.get("outlineLevel")
        if ol is not None:
            level = int(ol) if isinstance(ol, (int, str)) else level
        # Only include styles likely to be headings
        is_heading = (re.match(r'^Heading[1-9]$', sid)
                      or "标题" in name or "章" in name or "节" in name
                      or "条" in name or "heading" in name.lower()
                      or ol is not None)
        if is_heading and level < 3:
            heading_styles.append((name, level, s))
    heading_styles.sort(key=lambda x: x[1])

    # Deduplicate by name
    seen = set()
    heading_styles = [hs for hs in heading_styles
                      if not (hs[0] in seen or seen.add(hs[0]))]

    placeholders = []
    for name, level, sinfo in heading_styles[:2]:
        ctor = _style_heading_js(name, sinfo)
        placeholders.append(
            f'      new Paragraph({{ {ctor}, children: [new TextRun("{name}")] }})'
        )
    if not placeholders:
        placeholders.append(
            '      new Paragraph({ children: [new TextRun("Body text.")] })'
        )
    if len(placeholders) == 1:
        placeholders.append(
            '      new Paragraph({ children: [new TextRun("Body text.")] })'
        )

    # Determine if we need HeadingLevel import
    need_heading_level = any(
        "heading: HeadingLevel" in pl for pl in placeholders
    )

    import_line = "const { Document, Packer, Paragraph, TextRun, AlignmentType,"
    import_line += " HeadingLevel," if need_heading_level else ""
    import_line += "\n    PageOrientation, PageNumber } = require('docx');"

    return f"""{import_line}
const fs = require('fs');
const DF = {font_js};
const doc = new Document({{
  styles: {{ default: {{ document: {{ run: {{ font: DF, size: {ds}, color: "000000" }} }} }},
    paragraphStyles: [
{",\n".join(styles_js)}
    ] }},
  sections: [{{ properties: {{ page: {{ size: {{ width: {w}, height: {h} }},
    margin: {{ top: {mg.get('top',1440)}, right: {mg.get('right',1440)}, bottom: {mg.get('bottom',1440)}, left: {mg.get('left',1440)} }} }} }},
    children: [
{",\n".join(placeholders)}
    ] }}]
}});
Packer.toBuffer(doc).then(b => fs.writeFileSync(process.argv[2], b));"""


def _run_opts(run: dict, df, ds: int) -> list[str]:
    o = []
    if "font" in run:
        f = run["font"]
        if isinstance(f, dict):
            ascii = f.get("ascii","Times New Roman")
            if "eastAsia" in f:
                o.append(f'font: {{ ascii: "{ascii}", hAnsi: "{ascii}", eastAsia: "{f["eastAsia"]}", cs: "{ascii}" }}')
            else: o.append(f'font: "{ascii}"')
    if "size" in run:
        s = run["size"]; o.append(f"size: {int(s.get('val',ds)) if isinstance(s,dict) else int(s)}")
    if "bold" in run: o.append("bold: true")
    if "italic" in run: o.append("italics: true")
    if "color" in run:
        c = run["color"]; o.append(f'color: "{c.get("val","000000") if isinstance(c,dict) else c}"')
    else: o.append('color: "000000"')
    if "vertAlign" in run:
        v = run["vertAlign"]; vv = v.get("val","") if isinstance(v,dict) else v
        if vv == "superscript": o.append("superScript: true")
        elif vv == "subscript": o.append("subScript: true")
    if "underline" in run:
        u = run["underline"]; uv = u.get("val","single") if isinstance(u,dict) else u
        if uv != "none": o.append(f'underline: {{ type: "{uv}" }}')
    return o


def _para_opts(para: dict) -> list[str]:
    o = []
    if "spacing" in para:
        sp = para["spacing"]; parts = []
        for k in ("after","before","line"):
            if k in sp: parts.append(f"{k}: {int(sp[k])}")
        if "lineRule" in sp: parts.append(f'lineRule: "{sp["lineRule"]}"')
        if parts: o.append("spacing: { " + ", ".join(parts) + " }")
    if "align" in para:
        m = {"left":"LEFT","center":"CENTER","right":"RIGHT","both":"JUSTIFIED"}
        o.append(f"alignment: AlignmentType.{m.get(para['align'], 'LEFT')}")
    if "outlineLevel" in para: o.append(f"outlineLevel: {para['outlineLevel']}")
    return o if o else ["spacing: { after: 120 }"]


def _heuristic_outline_level(name: str) -> int:
    """Guess outline level (0-based) from style name.
    Returns 0 for unknown styles."""
    key = name.lower().strip()
    mapping = {
        "heading 1": 0, "heading 2": 1, "heading 3": 2,
        "heading 4": 3, "heading 5": 4, "heading 6": 5,
        "heading 7": 6, "heading 8": 7, "heading 9": 8,
        "title": 0, "subtitle": 1,
        "章标题": 0, "节标题": 1, "条标题": 2,
        "标题 1": 0, "标题 2": 1, "标题 3": 2,
        "标题 4": 3, "标题 5": 4, "标题 6": 5,
        "chapter": 0, "section": 1, "subsection": 2,
    }
    return mapping.get(key, 0)


def _style_heading_js(name: str, style_info: dict) -> str:
    """Return a docx-js Paragraph constructor fragment.

    Uses { heading: HeadingLevel.HEADING_N } when the style maps to a
    built-in heading.  Uses { style: "name", outlineLevel: N } for
    custom-named styles.
    """
    sid = style_info.get("styleId", "")
    para_info = style_info.get("paragraph", {})
    ol = para_info.get("outlineLevel")
    if ol is not None:
        level = int(ol) if isinstance(ol, (int, str)) else 0
    else:
        level = _heuristic_outline_level(name)
    is_builtin = bool(re.match(r'^Heading[1-9]$', sid))
    if is_builtin:
        hnum = int(sid[-1]) if sid[-1].isdigit() else 1
        return f"heading: HeadingLevel.HEADING_{hnum}"
    else:
        return f'style: "{name}", outlineLevel: {level}'


# ---- apply: existing document ----
def _apply_existing(tpl: dict, docx: Path, fix_formulas: bool = False) -> None:
    with tempfile.TemporaryDirectory() as td:
        up = Path(td) / "u"
        _run(str(SKILL / "scripts/office/unpack.py"), str(docx), str(up))
        _patch_styles(up, tpl)
        _patch_page(up, tpl)
        cmd = [str(SKILL / "scripts/office/pack.py"), str(up), str(docx),
               "--original", str(docx), "--validate", "false"]
        if fix_formulas:
            cmd.append("--fix-formulas")
        _run(*cmd)


def _patch_styles(up: Path, tpl: dict) -> None:
    sx = up / "word" / "styles.xml"
    if not sx.exists(): return
    regw(); tree = parse(sx); root = tree.getroot()

    if (dr := tpl.get("defaults",{}).get("run")):
        dd = root.find(wtag("docDefaults"))
        if dd is None: dd = ET.SubElement(root, wtag("docDefaults"))
        rpd = dd.find(wtag("rPrDefault"))
        if rpd is None: rpd = ET.SubElement(dd, wtag("rPrDefault"))
        rpr_old = rpd.find(wtag("rPr"))
        if rpr_old is not None: rpd.remove(rpr_old)
        rpr = ET.SubElement(rpd, wtag("rPr"))
        _merge_rpr(rpr, dr)

    existing_names = set()
    for style in root.iter(wtag("style")):
        nm = style.find(wtag("name"))
        if nm is not None: existing_names.add(nm.get(wtag("val")))

    for name, s in tpl.get("styles",{}).items():
        found = False
        for style in root.iter(wtag("style")):
            nm = style.find(wtag("name"))
            if nm is not None and nm.get(wtag("val")) == name:
                found = True
                if "run" in s:
                    rpr = style.find(wtag("rPr"))
                    if rpr is None: rpr = ET.SubElement(style, wtag("rPr"))
                    _merge_rpr(rpr, s["run"])
                if "paragraph" in s:
                    ppr = style.find(wtag("pPr"))
                    if ppr is None: ppr = ET.SubElement(style, wtag("pPr"))
                    _merge_ppr(ppr, s["paragraph"])
                break
        if not found and name not in existing_names:
            _add_style(root, name, s)
    write(tree, sx)


_TAG_JSON = {"font":"rFonts","size":"sz","bold":"b","italic":"i","color":"color",
             "vertAlign":"vertAlign","underline":"u","kern":"kern","sizeCs":"szCs","boldCs":"bCs"}

def _merge_rpr(rpr: ET.Element, props: dict) -> None:
    for jk, xt in _TAG_JSON.items():
        if jk not in props: continue
        v = props[jk]; old = rpr.find(wtag(xt))
        if old is not None: rpr.remove(old)
        if v in (False,"0","","none"): continue
        e = ET.SubElement(rpr, wtag(xt))
        if isinstance(v, dict):
            for kk, vv in v.items(): e.set(wtag(kk), str(vv))
        elif xt == "rFonts":
            if isinstance(v, dict):
                for a, vv in v.items(): e.set(wtag(a), str(vv))
            else:
                for a in ("ascii","hAnsi","eastAsia","cs"): e.set(wtag(a), str(v))
        elif xt in ("b","bCs","i","iCs"): e.set(wtag("val"), "1" if v in ("1",True) else str(v))
        else: e.set(wtag("val"), str(v))


def _merge_ppr(ppr: ET.Element, props: dict) -> None:
    for key, tag in [("spacing","spacing"),("align","jc"),("outlineLevel","outlineLvl")]:
        if key in props:
            old = ppr.find(wtag(tag))
            if old is not None: ppr.remove(old)
            e = ET.SubElement(ppr, wtag(tag))
            v = props[key]
            if isinstance(v, dict):
                for kk, vv in v.items(): e.set(wtag(kk), str(vv))
            else: e.set(wtag("val"), str(v))


def _add_style(root: ET.Element, name: str, s: dict) -> None:
    """Add a new style element to styles.xml."""
    stype = s.get("type","paragraph")
    sid = s.get("styleId","")
    elem = ET.SubElement(root, wtag("style"))
    elem.set(wtag("type"), stype)
    if sid: elem.set(wtag("styleId"), sid)
    if "customStyle" in s: elem.set(wtag("customStyle"), "1")
    nm = ET.SubElement(elem, wtag("name")); nm.set(wtag("val"), name)
    if bo := s.get("basedOn",""):
        elem.set(wtag("basedOn"), bo)
    if "run" in s:
        rpr = ET.SubElement(elem, wtag("rPr"))
        _merge_rpr(rpr, s["run"])
    if "paragraph" in s:
        ppr = ET.SubElement(elem, wtag("pPr"))
        _merge_ppr(ppr, s["paragraph"])


def _patch_page(up: Path, tpl: dict) -> None:
    pg = tpl.get("page",{})
    if not pg: return
    dx = up / "word" / "document.xml"
    if not dx.exists(): return
    regw(); tree = parse(dx); root = tree.getroot()
    sect = root.find(f"{wtag('body')}/{wtag('sectPr')}")
    if sect is None: return
    for json_key, tag, xml_attr in [("width","pgSz","w"),("height","pgSz","h"),
        ("orientation","pgSz","orient")]:
        if json_key in pg:
            e = sect.find(wtag(tag))
            if e is None: e = ET.SubElement(sect, wtag(tag))
            e.set(wtag(xml_attr), str(pg[json_key]))
    if "margins" in pg:
        e = sect.find(wtag("pgMar"))
        if e is None: e = ET.SubElement(sect, wtag("pgMar"))
        for a, v in pg["margins"].items(): e.set(wtag(a), str(v))
    write(tree, dx)


# ---- CLI ----
def main():
    p = argparse.ArgumentParser(description="Extract/apply docx formatting templates")
    sp = p.add_subparsers(dest="cmd", required=True)
    e = sp.add_parser("extract"); e.add_argument("template"); e.add_argument("-o","--output")
    a = sp.add_parser("apply"); a.add_argument("template"); a.add_argument("-o","--output")
    a.add_argument("--to")
    a.add_argument("--no-fix-formulas", action="store_false", dest="fix_formulas",
                   default=True,
                   help="Skip automatic _/^ marker conversion (enabled by default)")
    args = p.parse_args()
    if args.cmd == "extract":
        src = Path(args.template)
        if not src.exists(): print(f"error: {src} not found", file=sys.stderr); sys.exit(1)
        tpl = _extract(src)
        out = args.output or (src.stem + "_template.json")
        json.dump(tpl, open(out,"w",encoding="utf-8"), indent=2, ensure_ascii=False)
        print(f"Extracted → {out}")
    else:
        tpl = json.load(open(args.template, encoding="utf-8"))
        if args.to:
            target = Path(args.to)
            if not target.exists(): print(f"error: {target} not found", file=sys.stderr); sys.exit(1)
            _apply_existing(tpl, target, fix_formulas=args.fix_formulas)
            print(f"Applied template → {target}")
        else:
            out = args.output or "output.docx"
            script = _gen_js(tpl)
            with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False, encoding="utf-8") as f:
                f.write(script); spath = f.name
            out_path = Path(out).resolve()
            try:
                _run_node(spath, str(out_path))
                # If --fix-formulas, unpack → fix → repack the newly created docx
                if args.fix_formulas:
                    with tempfile.TemporaryDirectory() as td:
                        up = Path(td) / "u"
                        _run(str(SKILL / "scripts/office/unpack.py"), str(out_path), str(up))
                        from fix_formulas import fix_formulas as run_fix
                        modified = run_fix(up)
                        if modified:
                            print(f"  fix_formulas: repaired {modified} run(s)")
                        _run(str(SKILL / "scripts/office/pack.py"), str(up), str(out_path),
                             "--original", str(out_path), "--validate", "false")
                print(f"Created → {out_path}")
            finally:
                Path(spath).unlink(missing_ok=True)

if __name__ == "__main__":
    main()
