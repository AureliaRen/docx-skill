r"""Extract/apply .docx formatting templates, and fill {{token}} placeholders.

Usage:
    python docx_template.py extract template.docx -o tpl.json
    python docx_template.py apply tpl.json -o new.docx [--fix-formulas]
    python docx_template.py apply tpl.json --to existing.docx [-o out.docx] [--fix-formulas]
    python docx_template.py fill template.docx values.json output.docx [--strict]
    python docx_template.py template.docx values.json output.docx      # fill 简写

--fix-formulas     Opt-in second channel: after writing, convert LaTeX / linear
                   formulas into native Office Math OMML (runs docx_mathify).
                   Remote lineage; off by default.
--no-fix-formulas  Turn off the *default* marker channel (paired _xx_ / ^xx^ ->
                   vertAlign runs), which is on unless you say otherwise.

                   These two flags drive two separate channels on two separate
                   defaults (``fix_formulas`` / ``omml_formulas``), and they are
                   mutually exclusive on the command line: passing both is an
                   argparse error (exit 2) rather than a silent merge.

``fill`` 是 hermes 版 CLI 的位置参数形式：不写子命令、直接给三个路径时也走 fill
（``_with_default_subcommand()`` 自动补上），所以旧脚本不用改。

解决什么问题
------------
模板套用分两半：

* ``extract`` —— 把一份成品 docx 的页面设置、docDefaults、全部样式抽成 JSON
  （这是"照着这份排版再做一份"的取证步骤）；
* ``apply`` —— 两条路：``-o``（无 ``--to``）走 node + docx-js 直接生成一份带同样式的
  空骨架；``--to`` 把 JSON 打回任意现有 docx（合并 rPr/pPr、改页面设置）。
  ``--to`` 同时给 ``-o`` 时**写到 ``-o``、``--to`` 那份原文件一个字节都不动**；
  只给 ``--to`` 时原地改写，但会先写一份 ``<file>.bak`` 备份。

坑（从 source-snapshot 原版踩过来后改的架构）
--------------------------------------------
1. **原版靠 subprocess 调 ``scripts/office/unpack.py`` / ``pack.py``**——那两个
   脚本及其 validators 子包不入仓，subprocess 链一断整条 extract/apply 路径就死。
   本版改为**进程内 zipfile 直读直写**：extract 直接从 zip 里读
   ``word/styles.xml`` / ``word/document.xml``；apply 用"临时文件 + shutil.move"
   就地改写（同 fix_fonts v2 的模式）。zip 重写保持原条目顺序，新条目追加在尾。
2. **``--fix-formulas`` 原版 import 的是 ``scripts/office/fix_formulas.py``**，
   同样不入仓。本版内置一个最小实现：单个 ``w:t`` 内成对 ``_xx_`` → subscript、
   ``^xx^`` → superscript。只用**成对包裹制**——不碰孤立的下划线/尖号，避免把
   LaTeX 式 ``x_i`` 的裸下划线误伤。**注意名字漂移**：原版这个 flag 指的就是这条
   marker 通道；合并后 ``--fix-formulas`` 已改指 OMML 通道（``dest=omml_formulas``，
   opt-in 默认关），marker 通道改由默认值（开）+ ``--no-fix-formulas`` 控制。
3. apply ``--to`` 合并样式时，同名样式只改 rPr/pPr 里 JSON 给到的键，不整段
   覆盖，保住目标文档已有的其余格式。
4. ElementTree 没有父指针：要把一个 ``w:t`` 拆成多个 run 时不能就地改，
   得按 run 逐个重写再 ``parent.remove/insert`` 回去（走 ``_fix_formula_markers``
   的两段式遍历）。
"""
from __future__ import annotations

import argparse, json, re, shutil, subprocess, sys, tempfile, zipfile
import io
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parent))

import _bootstrap  # noqa: F401 - flat-layout bootstrap + minidom / UTF-8 fallbacks

from docx_utils import W, wtag, XML_SPACE, regw

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
        elif tag in _PPR_MAP: r[_PPR_MAP[tag]] = s
        else: r[tag] = s
    return r or None


# ---- subprocess wrappers ----
def _env(): e = __import__('os').environ.copy(); e.setdefault("PYTHONIOENCODING","utf-8"); return e


def _run_node(script: str, output: str) -> None:
    subprocess.run(["node", script, output], check=True,
                   capture_output=True, encoding="utf-8", errors="replace", env=_env())


# ---- zip IO（进程内，替代原版 subprocess unpack/pack）----
def _read_blobs(docx: Path) -> tuple[list[str], dict[str, bytes]]:
    with zipfile.ZipFile(docx) as z:
        names = z.namelist()
        blobs = {n: z.read(n) for n in names}
    return names, blobs


def _write_blobs(docx: Path, names: list[str], blobs: dict[str, bytes]) -> None:
    """就地改写：临时文件 → shutil.move。保持原条目顺序（新条目追加在尾）。"""
    tmp = docx.with_suffix(docx.suffix + ".tplfix.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            if n in blobs:
                zout.writestr(n, blobs[n])
        for n in blobs:
            if n not in names:
                zout.writestr(n, blobs[n])
    shutil.move(str(tmp), str(docx))


def _tree_to_bytes(tree: ET.ElementTree) -> bytes:
    """Serialize with the standard XML declaration Word expects."""
    buf = io.BytesIO()
    tree.write(buf, xml_declaration=True, encoding="UTF-8")
    return buf.getvalue()


# ---- extract ----
def _extract(docx: Path) -> dict:
    _, blobs = _read_blobs(docx)
    tpl: dict = {"page":{}, "defaults":{}, "styles":{}}

    sx = blobs.get("word/styles.xml")
    if sx is not None:
        regw(); tree = ET.ElementTree(ET.fromstring(sx))
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

    dx = blobs.get("word/document.xml")
    if dx is not None:
        regw(); tree = ET.ElementTree(ET.fromstring(dx))
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


# ---- inline formula-marker fix（替代原版 import 的 fix_formulas.py）----
def _split_markers(text: str):
    """Return [(seg_text, 'sub'|'sup'|None)] for one w:t text. Paired only."""
    out = []
    pos = 0
    for m in re.finditer(r"_([^_{}]+)_|\^([^\^{}]+)\^", text):
        if m.start() > pos:
            out.append((text[pos:m.start()], None))
        out.append((m.group(1) or m.group(2), "sub" if m.group(1) else "sup"))
        pos = m.end()
    if pos < len(text):
        out.append((text[pos:], None))
    return out


def _fix_formula_markers(doc: ET.Element) -> int:
    """Convert paired _xx_ / ^xx^ markers inside w:t to vertAlign runs.

    ElementTree 没有父指针，逐个 w:t 就地拆 run 做不到——改为按容器遍历，
    对每个容器下的每个 run 单独处理再回写。
    """
    fixed = 0
    for parent in doc.iter():
        runs = [c for c in list(parent) if c.tag == wtag("r")]
        for r in runs:
            ts = [c for c in r if c.tag == wtag("t")]
            if len(ts) != 1:
                continue
            t = ts[0]
            if not t.text or ("_" not in t.text and "^" not in t.text):
                continue
            segs = _split_markers(t.text)
            if all(kind is None for _, kind in segs):
                continue
            rpr = r.find(wtag("rPr"))
            new_runs = []
            for text, kind in segs:
                nr = ET.Element(wtag("r"))
                if rpr is not None:
                    nr.append(ET.Element(rpr.tag, dict(rpr.attrib)))
                if kind is not None:
                    e_rpr = nr.find(wtag("rPr"))
                    if e_rpr is None:
                        e_rpr = ET.SubElement(nr, wtag("rPr"))
                    va = ET.SubElement(e_rpr, wtag("vertAlign"))
                    va.set(wtag("val"), "subscript" if kind == "sub" else "superscript")
                    fixed += 1
                nt = ET.SubElement(nr, wtag("t"))
                nt.text = text
                if text and (text[0] == " " or text[-1] == " "):
                    nt.set(XML_SPACE, "preserve")
                new_runs.append(nr)
            idx = list(parent).index(r)
            parent.remove(r)
            for off, nr in enumerate(new_runs):
                parent.insert(idx + off, nr)
    return fixed


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

    # Join here, not inside the f-string. A backslash inside an f-string
    # expression part is PEP 701 syntax: Python 3.12 accepts it, 3.10 raises
    # "f-string expression part cannot include a backslash". The CI matrix
    # builds on 3.10, so the join has to happen before the literal.
    styles_block = ",\n".join(styles_js)
    placeholders_block = ",\n".join(placeholders)

    return f"""{import_line}
const fs = require('fs');
const DF = {font_js};
const doc = new Document({{
  styles: {{ default: {{ document: {{ run: {{ font: DF, size: {ds}, color: "000000" }} }} }},
    paragraphStyles: [
{styles_block}
    ] }},
  sections: [{{ properties: {{ page: {{ size: {{ width: {w}, height: {h} }},
    margin: {{ top: {mg.get('top',1440)}, right: {mg.get('right',1440)}, bottom: {mg.get('bottom',1440)}, left: {mg.get('left',1440)} }} }} }},
    children: [
{placeholders_block}
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
    names, blobs = _read_blobs(docx)
    if "word/styles.xml" in blobs:
        regw(); tree = ET.ElementTree(ET.fromstring(blobs["word/styles.xml"]))
        _patch_styles(tree.getroot(), tpl)
        blobs["word/styles.xml"] = _tree_to_bytes(tree)
    if "word/document.xml" in blobs:
        regw(); tree = ET.ElementTree(ET.fromstring(blobs["word/document.xml"]))
        _patch_page(tree.getroot(), tpl)
        if fix_formulas:
            n = _fix_formula_markers(tree.getroot())
            if n:
                print(f"  fix_formulas: repaired {n} run(s)")
        blobs["word/document.xml"] = _tree_to_bytes(tree)
    _write_blobs(docx, names, blobs)


def _patch_styles(root: ET.Element, tpl: dict) -> None:
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


def _patch_page(root: ET.Element, tpl: dict) -> None:
    pg = tpl.get("page",{})
    if not pg: return
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


# ---- --fix-formulas: native OMML conversion (remote template.py lineage) ----
def _convert_formulas(docx_path: Path) -> None:
    """Run the real formula chain over ``docx_path`` in place.

    Lazy-imports ``docx_mathify`` and hands it the file with ``--inplace``:
    linear formulas (LaTeX-ish) become native Office Math (OMML). This is a
    different mechanism from the default ``_fix_formula_markers`` pass
    (paired ``_xx_`` / ``^xx^`` -> vertAlign runs); the two coexist.
    """
    from docx_mathify import main as mathify_main
    argv = sys.argv
    try:
        sys.argv = ["docx_mathify.py", str(docx_path), "--inplace"]
        mathify_main()
    finally:
        sys.argv = argv


# ---- fill: {{token}} placeholders（对齐 hermes 版 CLI 行为）----
TOKEN_RE = re.compile(r"\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}")


def _fill(template: Path, values_path: Path, output: Path, strict: bool = False) -> dict:
    """Fill {{token}} in every paragraph (body + nested tables + header/footer).

    Token 的空白变体（``{{ name }}`` / ``{{name}}``）都填同一个值：先按
    ``TOKEN_RE`` 取出现过的规范名，再对该段里所有指向同一名字的**字面变体**
    逐个替换，所以两种写法不会漏。
    """
    from docx import Document            # lazy: 主路径 extract/apply 只用 ElementTree

    from docx_common import iter_all_paragraphs, replace_in_paragraph

    with open(values_path, encoding="utf-8") as f:
        values = json.load(f)
    if not isinstance(values, dict):
        raise ValueError(f"{values_path}: top-level JSON must be an object")

    doc = Document(str(template))
    filled: dict[str, int] = {}
    for para in iter_all_paragraphs(doc):
        for name in set(TOKEN_RE.findall(para.text)):
            if name not in values:
                continue
            for variant in set(
                    t.group(0) for t in TOKEN_RE.finditer(para.text)
                    if t.group(1) == name):
                n = replace_in_paragraph(para, variant, str(values[name]))
                filled[name] = filled.get(name, 0) + n

    remaining = sorted({m for para in iter_all_paragraphs(doc)
                        for m in TOKEN_RE.findall(para.text)})
    doc.save(str(output))
    return {"ok": not (strict and remaining), "output": str(output),
            "filled": filled, "unfilled_tokens": remaining}


# ---- CLI ----
_SUBCOMMANDS = ("extract", "apply", "fill")
_DOCX_SUFFIXES = (".docx", ".docm", ".dotx", ".dotm")


def _with_default_subcommand(argv=None) -> list[str]:
    """让 hermes 式位置参数调用 ``tpl.docx values.json out.docx`` 等价于 fill。

    argparse 的 subparsers 是 required=True，不写子命令会报 invalid choice；
    旧脚本（以及 harness 那套 e2e 测试）用的就是不写子命令的三位置参数形式，
    这里在解析前补上 ``fill``。

    只在**首参确实是 Word 文档路径**（扩展名命中 ``_DOCX_SUFFIXES``）时才补，
    其余情况原样交给 argparse —— 否则 ``docx_template.py bogus`` 会被误当成
    ``fill bogus``，报"缺 values/output"而不是"未知子命令"。
    """
    argv = list(sys.argv[1:] if argv is None else argv)
    if (argv and argv[0] not in _SUBCOMMANDS
            and not argv[0].startswith("-")
            and Path(argv[0]).suffix.lower() in _DOCX_SUFFIXES):
        argv.insert(0, "fill")
    return argv


def main(argv=None):
    p = argparse.ArgumentParser(
        description="Extract/apply docx formatting templates, fill {{token}} placeholders")
    sp = p.add_subparsers(dest="cmd", required=True)
    e = sp.add_parser("extract", help="dump page setup + styles to JSON")
    e.add_argument("template"); e.add_argument("-o","--output")
    a = sp.add_parser("apply", help="apply a template JSON to a new or existing docx")
    a.add_argument("template"); a.add_argument("-o","--output")
    a.add_argument("--to")
    # Two formula channels, two dests, two defaults:
    #   marker channel -> dest=fix_formulas,   default ON,  ``--no-fix-formulas``
    #   OMML channel   -> dest=omml_formulas,  default OFF, ``--fix-formulas``
    # The two names read like a +/- pair but they write different dests, so both
    # used to be accepted and took effect as an AND: the marker pass switched off
    # and docx_mathify then received raw ``_xx_`` markers. On a document that
    # carries markers that conversion fails, yet the CLI still exited 0
    # (measured 2026-11-11, P6/FIX-C5D). Defaults stay as they were; the pair is
    # now exclusive so "both" is a usage error instead of a silent merge.
    fx = a.add_mutually_exclusive_group()
    fx.add_argument("--no-fix-formulas", action="store_false", dest="fix_formulas",
                    default=True,
                    help="Turn off the default marker channel (paired _xx_ / ^xx^ "
                         "-> vertAlign runs); on unless you pass this")
    fx.add_argument("--fix-formulas", action="store_true", dest="omml_formulas",
                    default=False,
                    help="Opt-in second channel: LaTeX / linear formulas -> native "
                         "Office Math OMML (docx_mathify), independent of the "
                         "marker channel")
    f = sp.add_parser("fill", help="fill {{token}} placeholders from a JSON mapping")
    f.add_argument("template", help="input .docx containing {{tokens}}")
    f.add_argument("values", help="JSON file mapping token -> value")
    f.add_argument("output", help="output .docx path")
    f.add_argument("--strict", action="store_true",
                   help="exit 1 if any {{token}} remains unfilled")
    args = p.parse_args(_with_default_subcommand(argv))
    if args.cmd == "extract":
        src = Path(args.template)
        if not src.exists(): print(f"error: {src} not found", file=sys.stderr); sys.exit(1)
        tpl = _extract(src)
        out = args.output or (src.stem + "_template.json")
        json.dump(tpl, open(out,"w",encoding="utf-8"), indent=2, ensure_ascii=False)
        print(f"Extracted → {out}")
        return 0
    if args.cmd == "fill":
        tpl_path = Path(args.template)
        if not tpl_path.exists():
            print(f"error: {tpl_path} not found", file=sys.stderr); return 1
        values_path = Path(args.values)
        if not values_path.exists():
            print(f"error: {values_path} not found", file=sys.stderr); return 1
        from docx.opc.exceptions import PackageNotFoundError   # 只在 fill 分支需要
        try:
            result = _fill(tpl_path, values_path, Path(args.output), strict=args.strict)
        except json.JSONDecodeError as exc:
            print(f"error: {values_path}: invalid JSON: {exc}", file=sys.stderr); return 1
        except PackageNotFoundError as exc:      # 模板不是合法 docx 包
            print(f"error: {tpl_path} is not a .docx package: {exc}",
                  file=sys.stderr); return 1
        except (OSError, ValueError, KeyError, TypeError,
                zipfile.BadZipFile) as exc:
            print(f"error: fill failed: {exc}", file=sys.stderr); return 1
        print(json.dumps(result, ensure_ascii=False))
        return 0 if result["ok"] else 1
    else:
        tpl = json.load(open(args.template, encoding="utf-8"))
        if args.to:
            target = Path(args.to)
            if not target.exists(): print(f"error: {target} not found", file=sys.stderr); sys.exit(1)
            # P0-B: -o is the write path. Without this branch honouring it, the
            # flag was silently dropped and `apply tpl.json --to in.docx -o out.docx`
            # rewrote in.docx in place (no out.docx, no backup, exit 0).
            dest = Path(args.output) if args.output else target
            if dest.exists() and dest.resolve() == target.resolve():
                dest = target                     # -o pointing at --to: in place
            if dest is target:
                # In place: the original cannot survive, so snapshot it first.
                backup = Path(str(target) + ".bak")
                overwrote = backup.exists()
                shutil.copyfile(target, backup)
                _apply_existing(tpl, target, fix_formulas=args.fix_formulas)
                print(f"Applied template → {target}  (in place)")
                print(f"  backup → {backup}"
                      + ("  (previous backup overwritten)" if overwrote else ""))
                if args.omml_formulas:
                    _convert_formulas(target)
            else:
                if not dest.parent.is_dir():
                    print(f"error: output directory does not exist: {dest.parent}\n"
                          f"       {target} was NOT modified", file=sys.stderr)
                    return 1
                shutil.copyfile(target, dest)      # byte copy: source stays pristine
                _apply_existing(tpl, dest, fix_formulas=args.fix_formulas)
                print(f"Applied template → {dest}")
                print(f"  source untouched → {target}")
                if args.omml_formulas:
                    _convert_formulas(dest)
        else:
            out = args.output or "output.docx"
            script = _gen_js(tpl)
            with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False, encoding="utf-8") as f:
                f.write(script); spath = f.name
            out_path = Path(out).resolve()
            try:
                _run_node(spath, str(out_path))
                # The default marker channel (on unless --no-fix-formulas):
                if args.fix_formulas:
                    regw()
                    names, blobs = _read_blobs(out_path)
                    if "word/document.xml" in blobs:
                        tree = ET.ElementTree(ET.fromstring(blobs["word/document.xml"]))
                        n = _fix_formula_markers(tree.getroot())
                        if n:
                            print(f"  fix_formulas: repaired {n} run(s)")
                        blobs["word/document.xml"] = _tree_to_bytes(tree)
                        _write_blobs(out_path, names, blobs)
                if args.omml_formulas:
                    _convert_formulas(out_path)
                print(f"Created → {out_path}")
            finally:
                Path(spath).unlink(missing_ok=True)
        return 0


if __name__ == "__main__":
    sys.exit(main())
