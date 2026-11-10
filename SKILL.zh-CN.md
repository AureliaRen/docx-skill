---
name: docx
description: "纯离线 Word .docx 工具集：创建/编辑/读取/校验/批注/修订/模板套用，外加 LaTeX→OMML 原生公式链与表格几何三处一致。当用户要生成或改 Word 文档、批量替换文档里的公式、修表格列宽/字体、合并 docx 或导出 PDF 时使用。另有自动编号、题注与交叉引用、静态目录、分节页码四组能力，供需要交付即正确（打开就看得见、F9 前后一致）的正式文档使用。"
license: MIT
---

# docx — 离线 docx 处理

`scripts/` 全平铺，27 个文件：18 个 argparse CLI（含 `fix_zoom.py`）+ 1 个自检入口（`fix_phasor_dot.py`，无 argparse，`__main__` 跑回归用例）+ 7 个纯库模块（`docx_common.py` 带 `--api` 自述入口，没有 CLI 行为）+ 1 个 PowerShell 导出脚本。另有 `tools/word_accept.py` 做文档级外挂验收（Word COM 实测页码）。一律 `python -X utf8 scripts/<名>.py ...`，参数以各文件 `main()` 为准。

## 环境

需要 **Python 3.10+**，加上 `requirements.txt` 里的依赖：

```bash
python -m pip install -r requirements.txt
```

`latex2mathml` 只有公式功能才需要。OMML 生成另外要本机装 Microsoft Office 提供的
`MML2OMML.XSL`（转换器复用它那张样式表）——装在非常规位置时用 `MML2OMML_XSL` 指过去。

## 路由表

| 我要做什么 | 用哪个 |
|---|---|
| 从零生成文档（标题/段落/列表/表格/图片/页眉页脚） | `docx_create.py spec.json out.docx`（spec 格式见其 docstring） |
| 改现有文档（查找替换、改单元格、插删段落、套样式、TOC、页码） | `docx_edit.py <sub>` |
| 读内容 / 提结构 / 看样式 / 抽图 / 查修订 | `docx_read.py file.docx --text\|--structure\|--styles\|--images DIR\|--revisions` |
| 健康检查（包完整性、关系、图片 magic、样式引用） | `docx_validate.py file.docx`（`ok:false` → exit 1） |
| 批注增删查 | `docx_comments.py list\|add --target ... \|delete --id N` |
| 修订接受/拒绝 | `docx_revisions.py list\|accept-all\|reject-all\|accept --id N\|reject --id N` |
| `{{token}}` 占位符填充 | `docx_template.py fill tpl.docx values.json out.docx [--strict]` |
| 抽/套格式模板 | `docx_template.py extract\|apply` |
| 整篇公式批量转 OMML | `docx_mathify.py in.docx -o out.docx` |
| 插单条 LaTeX 公式 | `insert_formula.py unpacked/ "<latex>" --after "marker"` |
| 表格列宽 | `table_width.py report\|set\|sync\|auto <unpacked>/` |
| 字体统一 | `fix_fonts.py file.docx [--heading-threshold 14]` |
| 合并两份文档 | `merge_docs.py base.docx src.docx out.docx` |
| docx → PDF | `pwsh scripts/pdf_export.ps1 -DocxPath x.docx [-PdfPath y.pdf]` |
| `<w:zoom>` 缺 `w:percent` | `fix_zoom.py file.docx [--percent 100]` |
| 看编号定义 / 加定义 / 套编号 / 断号重编 / 去编号 | `docx_numbering.py list\|define\|apply\|restart\|detach` |
| 题注、交叉引用、书签、全量重算域 | `docx_crossref.py caption\|ref\|bookmark\|renumber\|list` |
| 写静态目录（真段落 + 点引线 + 实测页码） | `docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7` |
| 量段落落在第几页（目录页码的来源） | `python tools/word_accept.py in.docx --page-probe 1,9,43 --json out.json` |
| 分节：分节符 / 页码格式 / 页眉页脚 / 分栏 | `docx_section.py list\|break\|pagenum\|pagenums\|header\|footer\|cols` |

常用命令（均已对本机 `--help` 验过 exit=0）：

```bash
python scripts/docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex
python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display --after "marker"
python scripts/table_width.py set unpacked/ --table 1 --widths 2.4,4.2,2.4,5.6
python scripts/merge_docs.py base.docx src.docx merged.docx
python scripts/fix_fonts.py a.docx b.docx --heading-threshold 14
python scripts/docx_template.py fill tpl.docx values.json out.docx --strict
```

## 公式链：三步，缺一步就是半个公式

```
线性式 --(linear_to_latex)--> LaTeX --(latex2mathml)--> MathML --(MML2OMML.XSL via lxml)--> OMML
```

引擎在 `math_omml.py`：`latex_to_omml_lxml(latex, display=False, raw_spacing=False)` 返回带 `xmlns:m` 的元素，直接 append 进 `w:p` 即可；`assert_legal()` 会拒收两类让 Word 报「文件可能已损坏」的结构（`m:num/m:den/m:e/m:sub/m:sup/m:deg/m:fName/m:lim` 内嵌 `<m:oMath>`、`m:rad` 缺 `m:e`）。`MML2OMML.XSL` 随 Office 安装，`find_xslt()` 自动探测，`$env:MML2OMML_XSL` 可覆盖。公式字体默认 Times New Roman（`MATH_FONT` 覆盖；`"Cambria Math"` 还原 Word 观感，`""` 不注入）。

`linear_to_latex.py` 决定**什么算公式**。它刻意保守，好让正常行文活下来：`[0, 1]`、`f(x, y)`、中文小标题一律不动；`U=IR`、`ΣI=0` 才转。但「认出是一个式子」不等于「需要一个公式对象」——见下面的三阶分流。

### `$...$` 标记层：作者标注，不让正则猜

先把"哪里是公式"交给写文档的人：正文里用 `$...$` 包住 LaTeX。正则只负责切分，不负责判断语义。

```python
def _add_mixed(p, text):
    """解析 $...$ 标记为 OMML，其余为文本。作者知道哪里是公式，别让正则猜。"""
    pos = 0
    for m in re.finditer(r"\$([^$]+)\$", text):
        if m.start() > pos:
            _run(p, text[pos:m.start()])
        tex = m.group(1)
        try:
            el = latex_to_omml_lxml(tex)
            fix_phasor_dots(el)      # 相量点必调，否则点会浮大
            p._p.append(el)
        except Exception as exc:
            FAILS.append((tex, str(exc)[:150]))   # 记账，回退纯文本，绝不静默丢公式
            _run(p, tex)
        pos = m.end()
    if pos < len(text):
        _run(p, text[pos:])
```

`FAILS` 必须落盘或打印；静默丢公式是最难查的失败。

### 相量点（`\dot{}`）必须收尾

`latex2mathml` 把 `\dot{I}` 转成 `m:limUpp`，Word 里渲染成"大字 + 飘着的点"。`fix_phasor_dot.fix_phasor_dots(omml) -> int` 就地改成 `m:acc`（`\u0307` 组合点），返回改了几处。

分三层说清楚，别会错意：

1. **`latex_to_omml_lxml()` 本身不修** —— 它是纯转换函数，`\dot{I}` 出来仍是 `limUpp=1 / acc=0`。
2. **两个 CLI 已内置修复，正常走 CLI 不会出丑态** —— `docx_mathify.py` 的 `_to_omml()` 是所有转换的唯一出口、`insert_formula.py` 在 `_omml_stdlib()` 内转完即修；两者都计数，汇总行报「修正相量点 N」。
3. **只有手写调用方**（自己 `import math_omml` 直接转）才需手动补这一刀，好处是能顺手断言 `limUpp == 0` 自证。

实测基准：

```
İ=U/Z 成立                   → 修正相量点 1 / oMath=1 limUpp=0 acc=2 TNR=16
电流 İ_{L}+İ_{C} = İ 成立     → 修正相量点 3 / oMath=1 limUpp=0 acc=6
İ_{L}+İ_{C}（无运算符，不转） → oMath=0   ← 正常：find_math_spans 要求有运算符/等式
```

自检：`python scripts/fix_phasor_dot.py`。

### 三阶分流：认得出来，未必需要公式对象

`docx_mathify` 把每个识别到的 span 放进三阶之一。**判据顺序是固定的**：先看总开关，再看免排版阶，再看纯上下标阶，都不中才落 OMML。

| 阶 | 判据 | 输出 | 例 |
|---|---|---|---|
| 结构式 | 含 `_STRUCTURAL_CMDS` 里的命令（43 条：`\frac` `\sqrt` `\int` `\sum` `\begin{` `\left` `\overline` …） | OMML 公式对象 | `x^2/16+y^2/9=1`、`v=s/t` |
| 纯上下标 | 无结构命令，且至少有一个 `^` 或 `_` | 普通 run + `w:vertAlign` —— 真 Word 上下标，`Ctrl+Shift+=` 可还原全字号 | `a=9.8m/s2`、`10^-7` |
| 免排版 | 无结构命令、无 `^`/`_`、恰好一个 `=`、左侧单个 ASCII 字母、右侧数字加可选单位 | 普通文本 run | `s=10km`、`t=2h`、`m=5kg`、`v=3m/s` |

第三阶由 `docx_mathify.py` 里的 `no_layout_reason()` 判定。它那个单位尾巴是**白名单**，不是「任意字母」：每个因子都必须是 `linear_to_latex.UNITS` 的成员。所以 `y=2x`、`n=2t` 仍是公式——那是函数关系，不是「2 米」「2 吨」；`v=s/t` 也仍是真分式（`t` 故意不在 `UNITS` 里）。`x=3` 归第三阶，但记的是 `plain_equation` 而不是 `plain_quantity`。当前刻意保守的两处：`T=25℃`（`UNITS` 里没有 `℃`）与希腊字母量符号（`η=80%`）仍按公式出。

**纯上下标阶的代价**：落到这一阶的 span 会丢数学斜体，`s` 变正体。这是 2026-11-07 拍定的取舍，不是待修缺陷——凡是压根不需要二维排版的东西，本来就该是文本而不是公式对象；做成公式对象意味着改一个字符要开公式编辑器、粘到别处会掉形、行高还被公式引擎撑开。

要斜体就从**输入形状**给，不要从阶层给：把式子写成带上下标的形式（`a=9.8m/s2`，自然落进纯上下标阶），或者在 `docx_create` 的 run spec 里显式写 `vert_align`——`superscript` / `subscript` 布尔值可作别名，而显式 `vert_align` 优先于别名。

**两个开关**，都是 `docx_mathify.py` 的模块级常量：

| 常量 | 默认 | 设成 `False` 会怎样 |
|---|---|---|
| `MIXED_SUPERSCRIPT` | `True` | 全部 span 走 OMML —— 这是总开关 |
| `NO_LAYOUT_NEEDED` | `True` | 只停第三阶，纯上下标阶照旧生效 |

`NO_LAYOUT_NEEDED` 坐在 `MIXED_SUPERSCRIPT` **之内**：关掉总开关两者一起停，这才让「`MIXED_SUPERSCRIPT = False` 等价于旧行为」这句话成立。

## 公式之后：字体与表格

**`w:rPr` 两条铁律**（违反任一条 Word 就弹「文件可能已损坏」，`fix_fonts.py` 里叫 `dedupe_rpr()` / `set_fonts()`）：

1. 同名子元素只能有一个 —— `w:rPr` 是 `xsd:sequence`、多数 `maxOccurs=1`，先 dedupe 只留最后一个。
2. `w:rFonts` 必须是 `w:rPr` 的**第一个**子元素 —— `rpr.insert(0, rf)`，不是 `append`。

**表格几何**：`tblLayout` 缺失（autofit）时渲染器把 `tcW` 当宽度提示；`tblLayout=fixed` 时**只认 `tblGrid/gridCol`、忽略 `tcW`**。`table.autofit=False` 写的正是 fixed，所以只设 `cell.width` 会静默出等宽列。三处必须同时一致：`tblW` + `gridCol` + `tcW`（单位 dxa，1 cm = 566.93 dxa）。`table_width.py set/sync/auto` 干的就是这件事。**先 `table_width.py report` 看清现状再决定**：`sync` 是把现有 `tcW` 抄进 grid，`set --widths` 才是显式钉死；不看 report 就 `sync`，等于把错的 `tcW` 忠实复制进 grid。

## 编号与交叉引用

三支脚本分管这一块：

- `docx_numbering.py` —— `list` / `define` / `apply` / `detach` / `restart`
- `docx_crossref.py` —— `caption` / `ref` / `bookmark` / `renumber` / `list`
- `docx_toc.py` —— 静态目录：`build_static_toc` / `collect_outline` / `drop_existing_toc`

Word 的编号是**两级间接**，本工具链别的脚本从没暴露过这一层：`<w:abstractNum>` 定义格式（`lvlText` 之类，如 `[%1]`）、`<w:num>` 是与之挂钩的计数器实例、段落上的 `<w:numPr>` 指名 `numId`。**共享同一个 `numId` 的段落共享一个计数器**；要让一段列表重新从 1 起，就把它指到另一个 `<w:num>`。

顺带说清一个历史坑：`docx_create.py` 的 `numbered_list` 早先套的是 Word 的 `List Number` **样式**，那个样式不带 `numPr`，段落渲染出来根本没编号。现在是真定义 `abstractNum` + `num` 并挂真 `numId`。两个 block 传同一个 `num_id` 就接成一条序列，要断开就调 `docx_numbering.py restart`。

交叉引用是 `REF <bookmark>` 域，题注是包在书签里的 `SEQ <label> \* ARABIC` 域。本工具链写出的**每个域都带缓存结果**（`begin / instrText / separate / result / end`），所以编号在 Word 重算任何东西之前就已经看得见，F9 之后也仍然对。`renumber` 负责重算 `SEQ` 计数器并改写所有缓存的 `REF` 值——**这才是「读者从不更新域」也能读到正确编号的原因**。

## 静态目录：真段落，不是域

`docx_create.py` spec 里的 `toc` block 是一个**域**：条目在有人按 F9 之前并不存在，而手写进域里的缓存一按 F9 就被丢掉。当文件必须在**打开那一刻就是对的**（交付物、转 PDF、拿去做 diff），就别用域，用 `scripts/docx_toc.py` 把条目写成真段落。

```
python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7
```

```python
import docx_toc as toc
report = toc.build_static_toc(doc, levels=(1, 2, 3), pages=["1", "3", "7"],
                              title="目　录")
# report: entries / levels / styles / tab_pos_mm / tab_pos_source
#         dropped_entries / dropped_titles / adopted / section_index / inserted_at
```

每条目录项长成 `标题 + <w:tab/> + 页码`，套 `toc 1/2/3` 样式，在版心宽度处设一个右对齐、点引线的制表位。这个宽度由 `docx_create.content_width_mm(section)` 算出来，**绝不要硬编码**——六套测试卷的版心实测就分 170 mm 与 160 mm 两种。

**页码必须是量出来的，不是猜出来的。** `pages=None` 就把页码位留空；要填就用 `word_accept.py --page-probe` 实测，它报的 `paragraph_pages[].displayed_text` 是**读者实际看到的那个号**，取值来自 `Information(WD_ACTIVE_END_ADJUSTED_PAGE)`。注意它**不是物理页序**：一个 `lowerRoman start=1` 的节，会把 `i` 打在物理第 8 页上。

两遍法（先测量、再写目录）的完整做法与全部坑表见 `references/toc.md`。

重复构建是安全的、幂等的：每条目录项都挂 `_DocxTocEntry…` 书签，`drop_existing_toc` 只删**本模块自己写的**东西——用户自带的那份用内置 `toc 1` 样式的目录不会被动，除非显式传 `adopt_unmarked=True`。

`report` 里几个字段值得记住：`entries` / `levels` / `styles` 是实际落下的条目；`tab_pos_mm` 与 `tab_pos_source` 记录制表位位置及其来源（是量出来的还是推出来的）；`dropped_entries` / `dropped_titles` 是**被丢弃**的条目与标题——这两个非空就要查原因，别当噪音忽略；`adopted` 是认领的旧目录项；`section_index` / `inserted_at` 是插入位置。

## PDF 导出

```powershell
pwsh scripts/pdf_export.ps1 -DocxPath merged.docx -PdfPath merged.pdf
```

Word COM：`Documents.Open($path, $false, $true)`（第三参 ReadOnly，不占文件）+ `ExportAsFixedFormat($pdf, 17)`（17 = wdExportFormatPDF）+ `try/finally` 里 `Close($false)` / `Quit()`，防僵尸 WINWORD。**不要用 WPS 导出带 OMML 的文档**：文本层会残缺，公式区文字搜不到。

## 合并文档的三处裁定

- **sectPr 统一用 base**：src 的丢弃。否则两份页面设置不同的文档合并后，版心会静默变成 src 的。
- **numbering 按 `w:numId` 并集去重**：连带 `w:abstractNum` 依赖补齐；base 无 `numbering.xml` → 整搬 + `[Content_Types]` Override + rels 补条目。
- **styles 按 `w:styleId` 并集去重**：base 已有的不覆盖。

合并后自检：段落数 = base×2 + 1（分页符段）、`word/media/*` 数量守恒、`sectPr` 仅 1 份且逐属性等于 base。

## 库模块（无 CLI，只能 import）

| 模块 | 提供什么 |
|---|---|
| `linear_to_latex.py` | `find_math_spans` / `atoms_to_linear` / `linear_to_latex` / `repair_merged_scripts`、单位白名单 `UNITS` |
| `math_omml.py` | `latex_to_omml_lxml` / `latex_to_omml_string` / `assert_legal` / `apply_math_font` / `find_xslt` |
| `fix_phasor_dot.py` | `fix_phasor_dots(omml) -> int`（`__main__` 跑自检） |
| `docx_pack.py` | `unpack` / `pack`：解到目录、再打回 .docx；保留原条目顺序与压缩方式，往返一次不改内容即等价归档 |
| `docx_utils.py` | `wtag` / `register_all` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`、`FONT_HEITI`、`FONT_SONG` |
| `docx_common.py` | `iter_all_paragraphs` / `iter_part_roots` / `replace_in_paragraph`（`--api` 自述） |
| `encoding.py` | `_fix_stdio()`（Windows 下 stdout 切 UTF-8） |
| `minidom_compat.py` | `defusedxml.minidom` 缺失时兜底 stdlib |
| `_bootstrap.py` | 平铺自举：`sys.path` + minidom / UTF-8 兜底；**新脚本第一件事就是 `import _bootstrap`** |

两个 CLI（`table_width.py`、`insert_formula.py`）吃**解包目录**——把 .docx 解成一堆 XML 再直接改。`docx_template.py apply --to` 反过来直接就地改写打包好的 .docx，所以给它 `.docx` 文件本身，别给解包目录。解包/打回：

```python
from docx_pack import unpack, pack
unpack("in.docx", "unpacked/")
pack("unpacked/", "out.docx", original="in.docx")
```

`pack` 会沿用原文件的条目顺序与压缩方式，所以原样往返得到的是等价归档（字节级 diff 不会因为重新 deflate 而变噪音）。

`unpack` 把归档当**不可信输入**：条目名里的 `../`、绝对路径、反斜杠都会被拒绝，不让压缩包写到解包目录之外，命中即抛 `UnsafeArchiveError`。

## references/：细节在这里，别塞进本文件

| 文件 | 什么时候读 |
|---|---|
| `references/formula-guide.md` | 写公式链生产代码、调 `_add_mixed`、排查 Word 报损坏、处理 `%`/`℃`/`～`/绝对值/复合单位等源文本坑 |
| `references/table-geometry.md` | 表格列宽不生效、要决定 set 还是 sync、`w:zoom` 缺 `w:percent` |
| `references/font-scheme.md` | 中英文混排字体规范、`w:rPr` 顺序与去重、`--heading-threshold` 语义 |
| `references/merge-and-export.md` | 合并后版心/编号/图片出问题、Word COM 导出细节、PIL 多页 PDF 写法 |
| `references/revisions-and-comments.md` | 批注与修订的 OOXML 表示与命令行语义 |
| `references/toc.md` | 写静态目录：两遍法（先 `page-probe` 量页码、再写条目）、制表位与点引线、全套坑表 |

## 退出码

| 码 | 含义 |
|---|---|
| `0` | 成功 |
| `1` | 预期内的失败（文件缺失、marker 没命中、spec 非法） |
| `2` | 用法错误（argparse 拦下，或缺了必需的模式标志） |
| `3` | 部分成功 —— `docx_mathify` 把文档写出来了，但有公式转换失败 |

`docx_mathify` 返回 `3` 是**故意的**：输出文件有效、绝大多数公式已转，但至少一条公式还是线性文本。**交付前必须查退出码**，只看文件存不存在会把半个公式发出去。

## 约定（Conventions）

- 每个脚本自己把所在目录塞进 `sys.path`，所以任意工作目录都能跑 —— `python scripts/x.py` 与 `python x.py` 都行。
- 机器可读的脚本往 `stdout` 写 JSON，给人看的写散文。诊断信息与错误一律走 `stderr`。
- 退出码反映结果；解析输出之前先查退出码。
- 文件读写一律显式 UTF-8；Windows 上 stdio 由 `_bootstrap.py` 切成 UTF-8，不要再手写 reconfigure。
- `requirements.txt` **故意纯 ASCII**：pip 用系统 locale 读它，一个 em dash 就能让 CJK Windows 装不上。

## 硬规矩

- 报告/正式文档的公式一律走 LaTeX→OMML；只有 `H₂O` 级极简上下标才用 `vertAlign` run。
- 所有脚本都能 `python -X utf8` 直接跑；Windows 上 stdout 由 `_bootstrap` 切 UTF-8，不要再手写 reconfigure。
- 产物就地改写前一定先备份（`fix_fonts.py` / `docx_template.py apply --to` 用临时文件 + `shutil.move`）。
- 需要 node 的只有 `docx_template.py apply -o`（docx-js 骨架）；断链时用 `apply --to`。
