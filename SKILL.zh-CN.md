---
name: docx
description: "纯离线 Word .docx 工具集：创建/编辑/读取/校验/批注/修订/模板套用，外加 LaTeX→OMML 原生公式链与表格几何三处一致。当用户要生成或改 Word 文档、批量替换文档里的公式、修表格列宽/字体、合并 docx 或导出 PDF 时使用。另有自动编号、题注与交叉引用、静态目录、分节页码四组能力，供需要交付即正确（打开就看得见、F9 前后一致）的正式文档使用。"
license: MIT
---

# docx — 离线 docx 处理

`scripts/` 全平铺，**26 个 `.py` + 1 个 `.ps1` = 27 个文件**。**25 个是 CLI**（各自带 `if __name__ == "__main__"`
守卫，能直接从命令行跑），**`_bootstrap.py` 是唯一的纯库**；这 25 个里 **18 个有 argparse 表面**（`--help` 打印 `usage: …`），另 7 个有 `__main__` 但没有 argparse——6 个只打印一行「是库模块，import 它」自述（`docx_pack` / `docx_utils` / `encoding` / `linear_to_latex` / `math_omml` / `minidom_compat`），`fix_phasor_dot.py` 跑的是自检。另有 `tools/word_accept.py` 做文档级外挂验收（Word COM 实测页码，`--page-probe` 只属于它）。一律 `python -X utf8 scripts/<名>.py ...`，参数以各文件 `main()` 为准。

输出约定按脚本而异，这是刻意的：

- **stdout 给 JSON** —— `docx_create`、`docx_read`、`docx_edit`、`docx_validate`、`docx_comments`、`docx_revisions`、`docx_section`、`docx_numbering`、`docx_crossref`、`docx_toc`。这些拿去解析。
- **人类可读文本** —— `docx_mathify`、`table_width`、`insert_formula`、`docx_template`（`fill` 除外，它往 stdout 写 JSON）、`fix_fonts`、`fix_zoom`、`merge_docs`、`pdf_export`、`docx_pack`（裸跑会告诉你该怎么 import）、以及 `docx_common --api`（函数名 + 签名清单）。这些是给你看的进度。

退出码是可依赖的：`0` 成功、`1` 预期内的失败、`2` 用法错误。完整表与唯一的「部分成功」码 `3` 见[退出码](#退出码)。

## 环境

需要 **Python 3.10+**，加上 `requirements.txt` 里的依赖：

```bash
python -m pip install -r requirements.txt
```

`latex2mathml` 只有公式功能才需要。OMML 生成另外要本机装 Microsoft Office 提供的
`MML2OMML.XSL`（转换器复用它那张样式表）——装在非常规位置时用 `MML2OMML_XSL` 指过去。

**两项能力只在 Windows 上成立，而且是在调用时才炸、不在安装时预警。** **公式链**要
`MML2OMML.XSL`，而 `math_omml.find_xslt()` 只探测 Windows 的 Office 安装位置
（`%ProgramFiles%`、`%ProgramFiles(x86)%`、`%LOCALAPPDATA%`、写死的 `C:\Program Files…`），
Linux / macOS 上必抛 `FileNotFoundError`，除非自己放一份并设 `MML2OMML_XSL`。**PDF 导出**
走 Word COM。其余全部（创建 / 读取 / 编辑 / 校验 / 批注 / 修订 / 编号 / 交叉引用 / 静态目录 /
分节 / 表格几何 / 字体 / 合并 / 模板 / `docx_pack`）是纯 `python-docx` + `lxml`，不用装 Office。

## 在解包目录上工作

两个命令（`table_width.py`、`insert_formula.py`）吃**解包目录**——把 `.docx` 解成一堆 XML 再直接改。`docx_template.py apply --to`、`fix_fonts.py`、`fix_zoom.py` 相反：它们吃打包好的 `.docx`，经临时文件就地改写，所以给文件本身，别给解包目录。造一份解包目录：

```bash
python -c "import sys; sys.path.insert(0,'scripts'); from docx_pack import unpack; unpack('in.docx','unpacked/')"
# ...改 unpacked/ 下的文件...
python -c "import sys; sys.path.insert(0,'scripts'); from docx_pack import pack; pack('unpacked/','out.docx',original='in.docx')"
```

或者在你已有的 Python 里直接 import：

```python
from docx_pack import unpack, pack
unpack("in.docx", "unpacked/")
pack("unpacked/", "out.docx", original="in.docx")
```

`pack` 会沿用原文件的条目顺序与压缩方式，所以原样往返得到的是等价归档（字节级 diff 不会因为重新 deflate 而变噪音）。`original=` 给了一份缺失、空、或不是 zip 的原文件时会直接抛错，不会静默打出一个丢了继承信息的包。

`unpack` 把归档当**不可信输入**：凡是会落到解包目录之外的条目一律拒绝，抛 `UnsafeArchiveError`，且在写入任何字节之前就抛。被拒的形态如下，每一条都有测试覆盖：

- **`..` 穿越** —— `../escape.txt`，以及反斜杠写法 `..\..\evil.txt`（Windows 上 `zipfile` 层会先把它规整成 `../../evil.txt`，由 `normpath` 检查兜住）。
- **POSIX 绝对路径** —— `/abs/evil.txt`。
- **Windows 盘符** —— `D:/x.txt` 与 `D:x.txt`。两者都不是 POSIX 绝对路径，但 Windows 上 `pathlib` 认为它们带根，`dst / name` 会把 `dst` 整个丢掉。
- **UNC 共享路径** —— `\\server\share\x.txt`，按 drive prefix 拒。
- **Windows 保留设备名** —— `CON`、`NUL`、`AUX`、`PRN`、`CONIN$`、`CONOUT$`、`COM0`–`COM9`、`LPT0`–`LPT9`；只匹配最后一段、忽略扩展名、大小写不敏感，所以 `nul`、`CON.txt`、`word/AUX.xml` 全被拒。Win32 在任何目录下都会把这些名字解析成设备，而 `NUL` 甚至连写入都吞掉还报成功。
- **末段尾点或尾空格** —— `evil.txt ` 与 `evil3. .`。Win32 会把尾点尾空格剥掉，于是两条归档条目会撞成同一个文件名，解出来的树和归档对不上。

## 快查表

| 我要做什么 | 用哪个 |
|---|---|
| 从零生成文档 | `python scripts/docx_create.py spec.json out.docx` |
| 读正文 / 大纲 / 样式 | `python scripts/docx_read.py in.docx --text` |
| 抽图片 | `python scripts/docx_read.py in.docx --images DIR` |
| 看修订 | `python scripts/docx_read.py in.docx --revisions` |
| 查找替换 | `python scripts/docx_edit.py replace in.docx --find A --replace B -o out.docx` |
| 改表格单元格 | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X -o out.docx` |
| 插段落 | `python scripts/docx_edit.py insert in.docx --index 3 --text "..." -o out.docx` |
| 删段落 | `python scripts/docx_edit.py delete in.docx --index 3 -o out.docx` |
| 套段落样式 | `python scripts/docx_edit.py style in.docx --index 0 --style "Heading 1" -o out.docx` |
| 合并格式相同的相邻 run | `python scripts/docx_edit.py normalize in.docx -o out.docx` |
| 插目录**域** | `python scripts/docx_edit.py toc in.docx --index 1 -o out.docx` |
| 页脚加「第 X 页 / 共 Y 页」 | `python scripts/docx_edit.py page-numbers in.docx -o out.docx` |
| 健康检查（包完整性、关系、图片 magic、样式引用） | `python scripts/docx_validate.py in.docx` |
| 列批注 | `python scripts/docx_comments.py list in.docx` |
| 加批注 | `python scripts/docx_comments.py add in.docx --target "some text" --text "note"` |
| 删批注 | `python scripts/docx_comments.py delete in.docx --id 0`（**id 从 `0` 起**，`list` 会打印真实 id） |
| 修订 list / 全接受 / 全拒绝 | `python scripts/docx_revisions.py list\|accept-all\|reject-all in.docx` |
| `{{token}}` 占位符填充 | `python scripts/docx_template.py tpl.docx values.json out.docx` |
| 抽格式模板 | `python scripts/docx_template.py extract tpl.docx -o tpl.json` |
| 套格式模板 | `python scripts/docx_template.py apply tpl.json --to in.docx -o out.docx` |
| 诊断表格列宽 | `python scripts/table_width.py report unpacked/` |
| 钉死列宽 | `python scripts/table_width.py set unpacked/ --widths 2.4,4.2,4.2` |
| 把 `tcW` 抄进 grid | `python scripts/table_width.py sync unpacked/` |
| 按版心均分列宽 | `python scripts/table_width.py auto unpacked/ [--content-width 14.6]` |
| 整篇线性公式批量转 OMML | `python scripts/docx_mathify.py in.docx -o out.docx --report conv.md` |
| 插单条 LaTeX 公式 | `python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display` |
| 看编号定义 | `python scripts/docx_numbering.py list in.docx` |
| 加编号定义 | `python scripts/docx_numbering.py define in.docx --text "[%1]" -o out.docx` |
| 给段落套编号 | `python scripts/docx_numbering.py apply in.docx --index 155:171 --num-id 1` |
| 断号重编 | `python scripts/docx_numbering.py restart in.docx --index 7 --num-id 1` |
| 去编号 | `python scripts/docx_numbering.py detach in.docx --index 3:9` |
| 做带编号题注 | `python scripts/docx_crossref.py caption in.docx --index 42 --label 图 --text "架构"` |
| 插交叉引用 | `python scripts/docx_crossref.py ref in.docx --index 70 --target 图_1 --prefix "如图" --suffix 所示` |
| 给段落打书签 | `python scripts/docx_crossref.py bookmark in.docx --index 12 --name 图_1 -o out.docx` |
| 全量重算域 | `python scripts/docx_crossref.py renumber in.docx -o out.docx` |
| 列书签 / 域 | `python scripts/docx_crossref.py list in.docx` |
| 写静态目录 | `python scripts/docx_toc.py in.docx -o out.docx --title 目录 --pages 1,3,7`（`--pages` 一个条目一个值 —— 这里三个值表示源文档正好三条条目；个数不等就是 exit 2） |
| 分节：分节符 / 页码格式 / 页眉页脚 / 分栏 | `python scripts/docx_section.py list\|break\|pagenum\|pagenums\|header\|footer\|cols in.docx` |
| 量段落落在第几页 | `python tools/word_accept.py in.docx --page-probe 1,9,43 --json out.json` |
| 字体统一 | `python scripts/fix_fonts.py in.docx [more.docx ...]` |
| 合并两份文档 | `python scripts/merge_docs.py base.docx src.docx out.docx` |
| `<w:zoom>` 缺 `w:percent` | `python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]` |
| docx → PDF（文本层不掉） | `pwsh scripts/pdf_export.ps1 -DocxPath in.docx [-PdfPath out.pdf]` |

常用命令（均已对本机 `--help` 验过 exit=0）：

```bash
python scripts/docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex
python scripts/insert_formula.py unpacked/ "\frac{-b}{2a}" --display --after "marker"
python scripts/table_width.py set unpacked/ --table 1 --widths 2.4,4.2,2.4,5.6
python scripts/merge_docs.py base.docx src.docx merged.docx
python scripts/fix_fonts.py a.docx b.docx --heading-threshold 14
python scripts/docx_template.py fill tpl.docx values.json out.docx --strict
```

`docx_edit page-numbers` 是**幂等**的：页脚里已经有 `PAGE` 或 `NUMPAGES` 域（复杂域或 `<w:fldSimple>` 都算）时，原文件逐字节不动、就地重跑一个字节都不写；stdout 会多出 `changed` 与 `already_present` 两个键，脚本据此能分清「这次加的」和「本来就有」。

**`--index` 一律 0 基**，但它数什么、越界怎么办，各脚本不同：

| 脚本 | `--index` 数的是什么 | 越界 |
|---|---|---|
| `docx_edit insert` | body 段落；下标指新段落**插在谁之前**，所以合法区间是 `[0, N-1]` | exit `2`，报错里给合法区间 |
| `docx_edit delete` / `style` | body 段落 | exit `2`，报错里给合法区间 |
| `docx_edit toc` | body 段落 | exit `2`，报错里给合法区间 |
| `docx_crossref` | body 段落 | exit `2`，报错里给合法区间 |
| `docx_numbering` | **body + 表格**段落，按文档顺序；`lo:hi` 是半开区间，`--index` 可重复传 | exit `1`，报「selection matched no paragraphs」 |
| `docx_toc --insert-at` | body 段落 | exit `2`，报错里给合法区间 |

上表每一处拒绝都发生在**动文档之前**：不写文件、输入逐字节不变、stdout 保持空，报错走 stderr 一行 `error: …`，写清哪个选项、你给的值、以及什么区间才成立，例如 `error: delete --index 99 is out of range: legal range is [0, 3], current N=4`。负数下标同样按越界拒收，而不是交给 Python 的负索引去读：`-1` 在**任何持久意义上**都不是「最后一段」——增删一段，同一个参数就指向别处了。

`docx_toc` 报告里的 `inserted_at` 是另一回事：它是目录**实际**写入的位置，插入后量出来的，不是入参。

**`docx_section.py` 管分节级页面设置。** `list` 把每节几何打成 JSON；`break` 追加一节；`pagenum` 写 `<w:pgNumType>`（格式与起始值）；`pagenums` 往某一节的页脚写 PAGE/NUMPAGES 域；`header` / `footer` 写该节自己的页眉页脚文字；`cols` 设分栏数。除 `list` 外每个子命令都吃 `--section N`（默认 `-1`，即最后一节）。

```bash
python scripts/docx_section.py break in.docx --start-type new_page -o out.docx
python scripts/docx_section.py pagenum out.docx --section 1 --fmt lowerRoman --start 1 -o out.docx
python scripts/docx_section.py pagenums out.docx --template page_of -o out.docx
```

中间那行的 `lowerRoman start=1` 正是 `docx_toc` 提醒过的情形：该节会把 `i` 打在它物理起始的那一页上，所以**显示页码**与**物理页序**是两个不同的数。

**`docx_template.py apply` 带两条独立的公式通道、两套各自的默认值，命令行上互斥：**

| flag | 通道 | 默认 |
|---|---|---|
| `--no-fix-formulas` | marker 通道 —— 成对 `_xx_` / `^xx^` → `vertAlign` run | **开** |
| `--fix-formulas` | OMML 通道 —— LaTeX / 线性公式 → 原生 Office Math（内部跑 `docx_mathify`） | 关 |

这两个名字看着像一对 `+`/`-`，实际写的是两个不同的 dest（`fix_formulas` / `omml_formulas`），所以早先同传会一边关掉 marker 通道、一边把裸 `_xx_` 标记交给 OMML 转换器 —— 转换随之失败，而 CLI 仍退 `0`。现在是 argparse 冲突，exit `2`。不给 `-o` 时 `apply --to` 就地改写目标、但先写一份 `<文件>.bak`；给了 `-o`，`--to` 那份原文件一个字节都不动。

**`--fix-formulas` 会把 `docx_mathify` 的退出码原样往上报，包括 `3`。** OMML 通道是就地跑一遍 mathify，而那边的 `3` 意思是「文档已写出，但至少一条公式还是线性文本」—— 唯一一个「有产物 + 非零码」的出口。这里若仍旧报 `0`，就等于把半成品当完成品交给调用方。所以现在 `apply` 返回的就是 mathify 返回的那个值：`1` 和 `2` 原样透传，`3` 也透传为 `3`，只有 `0` 才归零。`--no-fix-formulas` 压根不调 mathify，行为与从前一致。

## 创建文档

`docx_create.py` 读一份 JSON spec。顶层键：`page`、`styles`、`header`、`footer`、`footer_page_numbers`、`blocks`。每个 block 有一个 `type`：

`footer_page_numbers`（顶层或 `section_break` 里）接受 `fmt`、`start`、`align`、`template`、`total`、`prefix`、`suffix`、`page_type`，以及 `font`、`size_pt`、`bold`——这三个作用于承载页码的 run：`font` 与 `size_pt` 文字 run 和域 run 都吃，`bold` 对承载页码数字的域 run 只在真值时生效（精确口径见下表）。它们早先一设就抛 `ValueError`，现在可用；中英混排的页脚里 `font` 会把四个 `w:rFonts` 槽（`ascii` / `hAnsi` / `eastAsia` / `cs`）一起绑死，`第 X 页 共 Y 页` 这种中文前缀才不会退回主题的东亚字体、让数字单独用你指定的那个。

**`bold` 作用在两类 run 上，而这两类走两套规则——所以 `"bold": false` 与不传 `bold` 不是同一个请求。** 一份 `page_of` 页脚由文字 run（`"Page "`、`" of "`、中文前后缀）和域 run（`PAGE` / `NUMPAGES` 域：页码数字本身，外加它的 begin / separate / end 三根骨架）混成。文字 run 走 `run.font.bold`，三态；域 run 走 `_rpr()`，那里写的是 `if bold:`，两态。拿同一份页脚、把 `font` 与 `size_pt` 按住不动，XML 里的差别就是这三行：

| `bold` | 文字 run | 域 run（页码数字） | `word/footer1.xml` | `<w:b` |
|---|---|---|---|---|
| `true` | `run.font.bold = True` → `<w:b/>` | 十个域 run 全写 `<w:b w:val="1"/>` **加** `<w:bCs w:val="1"/>` | 4232 字节 | 22 处：12 个 `w:b`（每个 run 一个）+ 10 个 `w:bCs` |
| `false` | `run.font.bold = False` → `<w:b w:val="0"/>`，显式**关** | **一个字节都不写**；`False` 与「没传」走同一分支 | 3912 字节 | 2 处，只在文字 run 上 |
| 不传 | 什么都不写，继续继承 | 什么都不写 | 3880 字节 | 0 处 |

这三个数就是全部结论：`false` 与不传**差 32 字节**（每个文字 run 16 字节），**两者不等价**；而且这两种写法**都不会给页码数字取消粗体**——数字在域 run 里，而域 run 只在 `bold` 为真值时才会拿到 `w:b`。所以如果页脚的 `Footer` 样式本身是粗体，`"bold": false` 只把 `"Page "` 与 `" of "` 压成不加粗，数字照旧粗着——看着像 bug，其实是本文写明的行为。要么在样式层关掉粗体，要么让 `bold` 保持真值、两类一起粗。（`section_break` 的 `header_bold` / `footer_bold` 是另一回事：那两个写的是纯文字 run、里面没有域，所以 `false` 是干净的显式关——实测唯一那个 run 上就是一条 `<w:b w:val="0"/>`——「不加粗」这个意思成立。）

| `type` | 字段 |
|---|---|
| `heading` | `text`、`level`（0–9） |
| `paragraph` | `text`，或 `runs`，另可选 `style` |
| `bullet_list` | `items` |
| `numbered_list` | `items`；可选 `num_id`（复用已有计数器）、`num_text`（默认 `%1.`）、`num_fmt`、`start`、`ilvl`、`indent_left`、`hanging` |
| `caption` | `label`（默认 `图`）、`text`、可选 `bookmark` —— 写入 `图 N  题注`，带 `SEQ` 计数器 + 书签 |
| `table_caption` | 字段同 `caption`，用于表格 |
| `crossref` | `target`（书签名），可选 `prefix`、`suffix`、`format`（`default`/`full`/`above-below`/`page`/`number`）、`cached`、`own_paragraph` |
| `table` | `header`、`rows`，可选 `style`（默认 `Table Grid`）、`header_bold` |
| `image` | `path`，可选 `width_mm` |
| `page_break` | — |
| `section_break` | `start_type`，可选 `page`、`header` / `footer`（及其 `_align` / `_font` / `_size_pt` / `_bold`）、`footer_page_numbers`（`fmt` / `start` / `align` / `template` / `total` / `prefix` / `suffix` / `page_type`，以及给页码 run 用的 `font` / `size_pt` / `bold`——见上面那张 `bold` 表：`false` 管得到文字 run，管不到域 run）、`different_first_page`、`cols`、`no_footer`、`reuse_last_paragraph`、`apply_to` |
| `toc` | 可选 `instr` —— 默认 `TOC \o "1-3" \h \z \u`；要图目录就传 `TOC \h \c "图"`。**这是一个域**：Word 必须更新（F9）之后才有条目 |

**spec 在写出任何一段之前先被校验；未知键与未知 block type 都是错误，不是 no-op。** 早先 `{"blockz": [...]}` 会造出一份合法但空白的文档、打印 `"blocks": 0`、退出 `0`，作者根本看不出内容没进去。现在一律 exit `2`，一行报错点名出问题的键，拼写接近时还给提示：

```
spec: unknown key "blockz" (did you mean "blocks"?); valid keys here: blocks, footer, footer_page_numbers, header, page, styles
blocks[0]: unknown block type 'paragraf' (did you mean "paragraph"?); choose one of: bullet_list, caption, crossref, heading, image, numbered_list, page_break, paragraph, section_break, table, table_caption, toc
```

exit `2` 的承诺是**什么都没写**，所以改完 spec 直接重试，不用猜是不是留了半份文档。值也不做钳制：超范围的值会连同它本该被钳进去的区间一起报出来。

`styles[]` 条目接受 `font`、`font_east_asia`、`size_pt`、`bold`、`italic`、`color`、`base`、`outline_level`——最后一个会写 `<w:pPr><w:outlineLvl>`，Word 才把该样式当标题。缺了它，模板自带的 `卷标题` / `ChapterTitle` 在 Word 眼里就是**正文**，长得再像也没用：导航窗格是空的，`TOC \o` 一条也收不到。

```json
{"styles": [
  {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
   "bold": true, "outline_level": 1}
]}
```

`outline_level` 是 1 基（`1` → `w:outlineLvl val="0"`）。`0` 表示「显式正文」（`val="9"`），给那种永远不该被收集的样式用。

`heading.level` 是 `0`–`9`，不是 `1`–`9`：python-docx 的 `add_heading` 把 `0` 当 Word 的 **Title** 样式、`1`–`9` 当 Heading 1–9。`10` 是 spec 错误，不会被钳。

## 静态目录：真段落，不是域

`toc` block 是一个**域**：条目在有人按 F9 之前并不存在，而手写进域里的缓存一按 F9 就被丢掉。当文件必须在**打开那一刻就是对的**（交付物、转 PDF、拿去做 diff），就别用域，用 `scripts/docx_toc.py` 把条目写成真段落。

```
# --pages 一个条目一个值：此行假设文档正好三条目录条目。
# 个数不匹配属用法错误（exit 2），不会被静默补空或丢弃。
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

值得记住的几个 flag：

| flag | 作用 |
|---|---|
| `--levels` | 收哪几级大纲，`1,2` 或 `1-3`（默认 `1,2,3`） |
| `--pages` | 每个条目一个页码；**个数与条目数不等就是用法错误（exit 2）** |
| `--title` | 目录本身的标题文字 |
| `--insert-at` | 插到这个 body 段落下标之前（默认：上次目录所在处，否则末尾）；越界是 exit `2` |
| `--section-index` | 用哪一节的版心宽度定右制表位 |
| `--indent-chars` | 每级左缩进字符数（默认 `0,2,4`） |
| `--adopt-unmarked` | 连本模块没写的旧目录一起替换——默认关，因为那可能是用户自己的 |
| `--dry-run` | 只报会写什么，不落盘 |

**页码必须是量出来的，不是猜出来的。** `pages=None` 就把页码位留空；要填就用 `tools/word_accept.py --page-probe` 实测，它报的 `paragraph_pages[].displayed_text` 是**读者实际看到的那个号**，取值来自 `Information(WD_ACTIVE_END_ADJUSTED_PAGE)`。`docx_toc.py` 自己没有测量用的 flag：你用 `word_accept.py` 量，再把数字通过 `--pages` 传回来。注意它**不是物理页序**：一个 `lowerRoman start=1` 的节，会把 `i` 打在物理第 8 页上。两遍法（先测量、再写目录）的完整做法与全部坑表见 `references/toc.md`。

重复构建是安全的、幂等的：每条目录项都挂 `_DocxTocEntry…` 书签，`drop_existing_toc` 只删**本模块自己写的**东西——用户自带的那份用内置 `toc 1` 样式的目录不会被动，除非显式传 `adopt_unmarked=True`。

`report` 里几个字段值得记住：`entries` / `levels` / `styles` 是实际落下的条目；`tab_pos_mm` 与 `tab_pos_source` 记录制表位位置及其来源（是量出来的还是推出来的）；`dropped_entries` / `dropped_titles` 是**被丢弃**的条目与标题——这两个非空就要查原因，别当噪音忽略；`adopted` 是认领的旧目录项；`section_index` / `inserted_at` 是插入位置。

`numbered_list` 现在是真的编号了。早先套的是 Word 的 `List Number` **样式**，那个样式不带 `numPr`，段落渲染出来根本没编号；现在是真定义 `abstractNum` + `num` 并挂真 `numId`。两个 block 传同一个 `num_id` 就接成一条序列，要断开就调 `docx_numbering.py restart`。

```json
{
  "blocks": [
    {"type": "heading",   "text": "Report title", "level": 1},
    {"type": "paragraph", "text": "Plain paragraph."},
    {"type": "paragraph", "runs": [
        {"text": "bold ", "bold": true},
        {"text": "italic", "italic": true},
        {"text": ", underlined", "underline": true}
    ]},
    {"type": "bullet_list", "items": ["first", "second"]},
    {"type": "numbered_list", "items": ["甲", "乙"], "num_text": "[%1]"},
    {"type": "caption", "label": "图", "text": "系统架构示意图"},
    {"type": "paragraph", "text": "系统架构如"},
    {"type": "crossref", "target": "图_1", "suffix": "所示。"},
    {"type": "table",  "header": ["Item", "Value"],
     "rows": [["Voltage", "220 V"]]}
  ]
}
```

`crossref` 把它的域追加到**上一段**，所以上面两个 block 能把一句话拆开写，读起来是 `系统架构如1所示。`。要它自己起一段就设 `"own_paragraph": true`。

## 编号与交叉引用

三支脚本分管这一块：

- `docx_numbering.py` —— `list` / `define` / `apply` / `detach` / `restart`
- `docx_crossref.py` —— `caption` / `ref` / `bookmark` / `renumber` / `list`
- `docx_toc.py` —— 静态目录：`build_static_toc` / `collect_outline` / `drop_existing_toc`

Word 的编号是**两级间接**，本工具链别的脚本从没暴露过这一层：`<w:abstractNum>` 定义格式（`lvlText` 之类，如 `[%1]`）、`<w:num>` 是与之挂钩的计数器实例、段落上的 `<w:numPr>` 指名 `numId`。**共享同一个 `numId` 的段落共享一个计数器**；要让一段列表重新从 1 起，就把它指到另一个 `<w:num>`。

交叉引用是 `REF <bookmark>` 域，题注是包在书签里的 `SEQ <label> \* ARABIC` 域。本工具链写出的**每个域都带缓存结果**（`begin / instrText / separate / result / end`），所以编号在 Word 重算任何东西之前就已经看得见，F9 之后也仍然对。`renumber` 负责重算 `SEQ` 计数器并改写所有缓存的 `REF` 值——**这才是「读者从不更新域」也能读到正确编号的原因**。

`docx_crossref bookmark` 是 `crossref` block 的先决条件：它给某一段打上具名书签，后面的 `ref`（或 spec 里的 `crossref` block）才有东西可指。名字必须唯一，同名再跑一次 exit `1`，报 `bookmark already exists`。

## 公式转换（LaTeX → 原生 Office Math）

两个入口共用一个引擎（`math_omml.py`）：

- **`docx_mathify.py` —— 批量。** 扫全篇，把写成 Unicode 字符 + 真上下标 run 的公式（`U`+sub`oc`、`√2`、`1kΩ=10³Ω`、`∠53.1°`）转成原生 OMML，中文行文、表格结构、字号都不动。`--dry-run` 只报不写；`--report md` 记下每对 before/after，另给计数与「哪些片段走了免排版阶」的逐段表。`--inplace` 会先写一份 `<名>.bak.docx`；`--inplace` 与 `-o` 同传是用法错误（exit `2`），因为落点无法判定；`-o` 与 `--dry-run` 同传同样是 exit `2` —— dry-run 一个文件都不写，`-o` 指定的落点永远不可能产生，原先却会静默并回显一个根本没被创建的路径。
- **`insert_formula.py` —— 单条。** 吃 LaTeX，把一条公式插进解包目录的锚点上。

`insert_formula.py` 有三条规矩，每一条都是为了让参数冲突当场报错、而不是静默走偏：

- **带锚点插入要报落点。** 写完之后脚本在树上反查新段落，把**实测坐标**追加到 stdout —— body 级锚点是 `at body[3]`，落进表格单元格的是 `at tc[1]`。下标是**容器自身直接子元素的下标**（含 `w:sectPr`、`w:tbl`），所以能直接和 `<w:body>` 的子元素序列对齐核对；它不是你在命令行里传的那个数。
- **`--dump` 与 `--after` / `--before` 同传是用法错误，exit `2`。** `--dump` 只打印 OMML、从不写文档，锚点会被无声吞掉。`--after` 与 `--before` 之间同理互斥。
- **空 marker 是用法错误，exit `2`。** `--after ""` 会一边看着像打了锚点、一边把公式追加到文档末尾；真想追加到末尾就别写这个 flag。

marker 一个都没命中时 exit `1`，报错里给出实际可搜范围 —— body 段落的 0 基区间加单元格数。

`linear_to_latex.py` 决定**什么算公式**。它刻意保守，好让正常行文活下来：`[0, 1]`、`f(x, y)`、中文小标题一律不动；`U=IR`、`ΣI=0` 才转。但「认出是一个式子」不等于「需要一个公式对象」——见下面的三阶分流。

公式链是三步，缺一步就是半个公式：

```
线性式 --(linear_to_latex)--> LaTeX --(latex2mathml)--> MathML --(MML2OMML.XSL via lxml)--> OMML
```

`latex_to_omml_lxml(latex, display=False, raw_spacing=False)` 返回带 `xmlns:m` 的元素，直接 append 进 `w:p` 即可；`assert_legal()` 会拒收两类让 Word 报「文件可能已损坏」的结构（`m:num/m:den/m:e/m:sub/m:sup/m:deg/m:fName/m:lim` 内嵌 `<m:oMath>`、`m:rad` 缺 `m:e`）。

### 三阶分流：认得出来，未必需要公式对象

`docx_mathify` 把每个识别到的 span 放进三阶之一。**判据顺序是固定的**：先看总开关，再看免排版阶，再看纯上下标阶，都不中才落 OMML。

| 阶 | 判据 | 输出 | 例 |
|---|---|---|---|
| 结构式 | 含 `_STRUCTURAL_CMDS` 里的命令（43 条：`\frac` `\sqrt` `\int` `\sum` `\begin{` `\left` `\overline` …） | OMML 公式对象 | `x^2/16+y^2/9=1`、`v=s/t` |
| 纯上下标 | 无结构命令，且至少有一个 `^` 或 `_` | 普通 run + `w:vertAlign` —— 真 Word 上下标，`Ctrl+Shift+=` 可还原全字号 | `a=9.8m/s2`、`10^-7` |
| 免排版 | 无结构命令、无 `^`/`_`、恰好一个 `=`、左侧单个 ASCII 字母、右侧数字加可选单位 | 普通文本 run | `s=10km`、`t=2h`、`m=5kg`、`v=3m/s` |

第三阶由 `docx_mathify.py` 里的 `no_layout_reason()` 判定。它那个单位尾巴是**白名单**，不是「任意字母」：每个因子都必须是 `linear_to_latex.UNITS` 的成员。所以 `y=2x`、`n=2t` 仍是公式——那是函数关系，不是「2 米」「2 吨」；`v=s/t` 也仍是真分式（`t` 故意不在 `UNITS` 里）。`x=3` 归第三阶，但记的是 `plain_equation` 而不是 `plain_quantity`。当前刻意保守的两处：`T=25℃`（`UNITS` 里没有 `℃`）与希腊字母量符号（`η=80%`）仍按公式出。

### 纯上下标阶的代价

落到这一阶的 span 会丢数学斜体，`s` 变正体。这是 2026-11-07 拍定的取舍，不是待修缺陷——凡是压根不需要二维排版的东西，本来就该是文本而不是公式对象；做成公式对象意味着改一个字符要开公式编辑器、粘到别处会掉形、行高还被公式引擎撑开。

要斜体就从**输入形状**给，不要从阶层给：把式子写成带上下标的形式（`a=9.8m/s2`，自然落进纯上下标阶），或者在 `docx_create` 的 run spec 里显式写 `vert_align`——`superscript` / `subscript` 布尔值可作别名，而显式 `vert_align` 优先于别名。

### 关掉分阶

两个开关，都是 `docx_mathify.py` 的模块级常量：

| 常量 | 默认 | 设成 `False` 会怎样 |
|---|---|---|
| `MIXED_SUPERSCRIPT` | `True` | 全部 span 走 OMML —— 这是总开关 |
| `NO_LAYOUT_NEEDED` | `True` | 只停第三阶，纯上下标阶照旧生效 |

`NO_LAYOUT_NEEDED` 坐在 `MIXED_SUPERSCRIPT` **之内**：关掉总开关两者一起停，这才让「`MIXED_SUPERSCRIPT = False` 等价于旧行为」这句话成立。

与这三阶同时修掉的两处源文本缺陷（连续 Unicode 上标、指数后的单位被读成分式），都写在 [formula-guide](references/formula-guide.md#two-defects-fixed-in-this-round) 里。

公式字体默认 **Times New Roman**。用环境变量 `MATH_FONT` 覆盖（`MATH_FONT="Cambria Math"` 还原 Word 默认观感；`MATH_FONT=""` 完全不注入）。

### `$...$` 标记层：作者标注，不让正则猜

先把「哪里是公式」交给写文档的人：正文里用 `$...$` 包住 LaTeX。正则只负责切分，不负责判断语义。

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

`latex2mathml` 把 `\dot{I}` 转成 `m:limUpp`，Word 里渲染成「大字 + 飘着的点」。`fix_phasor_dot.fix_phasor_dots(omml) -> int` 就地改成 `m:acc`（`\u0307` 组合点），返回改了几处。

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

## 表格几何

Word 算列宽的方式会让人踩空：

- `tblLayout` 缺失（autofit）时，每个单元格的 `<w:tcW>` 只是**提示**，列宽按它成比例分。
- `tblLayout type="fixed"` 时，渲染器**只认 `<w:tblGrid><w:gridCol>`、完全忽略 `tcW`**。默认的 grid 会让每列等宽，你怎么设都没用。

`table.autofit = False` 写的正是那个 `fixed`，所以只设 `cell.width` 会静默出等宽列。三处必须同时一致：`tblW` + `gridCol` + `tcW`，单位 dxa（1 cm = 566.93 dxa）。

`table_width.py` 干的就是这件事。**先 `table_width.py report` 看清现状再决定**：`sync` 是把现有 `tcW` 抄进 grid，`set --widths` 才是显式钉死，`auto` 则按可用版心均分（`--content-width`，默认 14.6 cm = A4 减去 3.17 cm 页边距）。

不看 report 就 `sync` 是那个坑：它会忠实地把当前的 `tcW`——错的也照抄——复制进 grid。而 `--widths` 的个数多于表格列数是用法错误、exit `2`，多出来的值它不替你丢：

```
error: table 1: --widths has 5 values but the table has 3 columns; refusing to drop the extra 2
```

## 字体：Word 强制的两条铁律

`fix_fonts.py` 一次统一三层——样式层、run 层、手动大标题层——并把中文字体绑到 `eastAsia` 槽，`ascii` / `hAnsi` / `cs` 保持 `Times New Roman`。之所以必须手改 XML，就是因为下面两条规矩。破哪一条，Word 都弹「文件可能已损坏」，不警告、也不静默重排。

1. **`w:rPr` 内同名子元素重复就是非法。** `w:rPr` 是 `xsd:sequence`、多数子元素声明 `maxOccurs="1"`，所以第二个 `w:rFonts` 就是 schema 违规——尽管解析器照样读得出来。真 Word 与 WPS 打开带重复的文件都会弹修复提示。`dedupe_rpr(rpr)` 在任何事之前先把同名子元素压到**最后一个**，于是后写的属性胜出——与宽容的读取器采用的优先级一致。
2. **`w:rFonts` 必须是 `w:rPr` 的第一个子元素。** 它开的是 `CT_RPr` 序列。删掉旧 `rFonts` 之后，新元素必须用 `rpr.insert(0, rf)` 塞回去；用 `append` 会把它停在末尾——那是一棵合法 XML 树、一个非法 `w:rPr`，Word 直接拒。那一行就是 `set_fonts(rpr, ea)`。

```
python scripts/fix_fonts.py in.docx [more.docx ...] [--heading-threshold 14]
```

run 按所在段落是不是标题分派；一个没有 `w:pStyle`、字号又达到 `--heading-threshold`（默认 14 pt）的 run 会被当成**手动大标题**——这是唯一能兜住「作者靠手调字号冒充标题、没套 `Heading 1`」的办法。`m:oMath` 里的公式 run 是例外：只改 `ascii` / `hAnsi` / `cs`，`eastAsia` 不动，因为 OMML 自己管字形选择，覆盖 `eastAsia` 会让公式观感肉眼可见地变差。

文件经临时文件 + `shutil.move` 就地改写，中断不会毁原文件。**但它不留额外副本**：与 `docx_mathify --inplace`、`docx_template apply --to` 不同，`fix_fonts` 不产 `.bak`，原件重要就自己先复制。

## `<w:zoom>` 与它必需的属性

python-docx 存盘时会写 `<w:zoom w:val="bestFit"/>`，却没写 `CT_Zoom` 声明为 `use="required"` 的 `w:percent`。它是良构 XML、非法 OOXML，严格的校验器会停在这里：

```
Element '{...}zoom': The attribute '{...}percent' is required but missing.
```

Word 容忍它，XSD 不容忍，任何按 schema 校验的流水线都会在这份「本工具链刚产出」的文档上失败。

`docx_create.py` 自己补这一刀，就在 `doc.save()` 之后：它调 `fix_zoom.fix_docx()`，所以它产的文档已经带 `w:percent="100"`，不用再手动跑第二遍。CLI 是给**别处来的**文档用的——你自己脚本里的 python-docx、旧文件、别的工具写的模板：

```
python scripts/fix_zoom.py in.docx [-o out.docx] [--percent 120]
```

有三件事脚本做对了，随手打的补丁做不到：

- 已经带 `w:percent` 的 `<w:zoom>` 原样返回，不碰。再加一个 `percent` 就是重复属性——畸形 XML，比缺属性更糟。它幂等：第二次跑只报 `zoom element already valid`。
- 重写 zip 时复用原来的 `ZipInfo` 条目，条目顺序与压缩方式都活下来。
- 包里没有 `word/settings.xml` 是合法的，不是坏了：脚本报 `nothing to do`、exit `0`，不去凭空造这个 part。

纯函数入口是 `fix_settings_xml(xml, percent) -> tuple[str, int]`：进 XML、出 `(patched_xml, sites_changed)`，不碰文件 I/O——测试就是直接 import 它、不 shell 出去。

## 合并文档

`merge_docs.py base.docx src.docx out.docx` 把 `src` 追加到 `base`，样式、编号、媒体都保住。三处裁定是刻意的，看结果之前值得先知道：

- **页面设置只取 `base`，`src` 的 `sectPr` 丢掉。** 两份都留不会报错——Word 取最后一节的几何，于是把 170 mm 版心的文档并进 160 mm 的那份，会静默把整个基底文档重排成错的版心。
- **numbering 按 `w:numId` 并集去重**，连带这些 id 依赖的 `w:abstractNum` 一起。`base` 完全没有 `numbering.xml` 时整份搬过去，外加 `[Content_Types]` Override 与一条 rels 条目——那才是让这个 part 真能被找到的东西。
- **styles 按 `w:styleId` 并集去重**；`base` 里已有的样式永不覆盖，所以基底文档的观感赢。

合并后自检三个数：段落数 = `base × 2 + 1`（多出来那段承载分页符）、`word/media/*` 数量守恒、`sectPr` 恰好剩 1 份且逐属性等于 `base` 的。合并结果看着不对，这三个数比读 XML 更快定位问题。

## PDF 导出（附加能力）

这不是本工具链的主业，frontmatter 里也这么写了——它是给**本工具链产出的**文档留的一条交付通道，不是通用转换器。

它只为一件具体的事存在：**文档含 OMML 公式。** WPS 导出这类文档时文本层是残的——公式区变成图片或直接消失，公式周围的文字在 PDF 里搜不到。Word 自己的导出器能把文本层留住，所以第二个前提是**装了 Microsoft Word 的 Windows**。两个条件同时成立，否则这脚本不是该用的工具。

```powershell
pwsh scripts/pdf_export.ps1 -DocxPath in.docx [-PdfPath out.pdf]
```

它走 Word COM：`Documents.Open($path, $false, $true)`（第三参 ReadOnly，源文件不被占用、也不沾「最近打开」状态），再 `ExportAsFixedFormat($pdf, 17)`（`17` = `wdExportFormatPDF`）。整段包在 `try` / `finally` 里，配 `Close($false)`（不保存）与 `Quit()`——因为 COM 异常若没走 `Quit()`，会留一个僵尸 `WINWORD` 进程攥着文件句柄，而「目录被锁」比一条错误信息难查得多。路径传绝对路径；COM 解析相对路径用的是进程工作目录，不是你的 shell 目录。

带 OMML 的文档别走 WPS。

## 库模块（只能 import）

**`_bootstrap.py` 是唯一的纯库**：没有 `__main__`、没有 argparse，每个脚本都在解析参数之前先 import 它——**新脚本第一件事就是 `import _bootstrap`**，一行覆盖 `sys.path` 自举、minidom 兜底与 UTF-8 stdio。下面七个模块同样以 import 为用途——其中六个带一行 `__main__` 自述、`fix_phasor_dot` 带自检，裸跑都不干活。它们一律按裸模块名 reach，因为每个脚本都把自己所在目录塞进 `sys.path`（见[约定](#约定conventions)）。

| 模块 | 提供什么 |
|---|---|
| `linear_to_latex.py` | `find_math_spans` / `atoms_to_linear` / `linear_to_latex` / `repair_merged_scripts`、单位白名单 `UNITS` |
| `math_omml.py` | `latex_to_omml_lxml` / `latex_to_omml_string` / `assert_legal` / `apply_math_font` / `find_xslt` |
| `fix_phasor_dot.py` | `fix_phasor_dots(omml) -> int`（`__main__` 跑自检） |
| `docx_pack.py` | `unpack` / `pack`：解到目录、再打回 .docx，见[在解包目录上工作](#在解包目录上工作) |
| `docx_utils.py` | `wtag` / `register_all` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`、`FONT_HEITI`、`FONT_SONG` |
| `encoding.py` | `_fix_stdio()`（Windows 下 stdout 切 UTF-8） |
| `minidom_compat.py` | `defusedxml.minidom` 缺失时兜底 stdlib |

`docx_common.py` **故意不在这张表里**：它有真 argparse 表面（`--api` 打印它导出的辅助函数），按判据算 CLI。

`minidom_compat` 比看上去重要。`defusedxml` 给 XML 解析加了防实体膨胀的硬化，但随包而来的解释器——DSH runtime、便携版 Python、精简容器——往往一个第三方包都没有，于是每个碰 docx XML 的脚本都在 **import 阶段**就死。这个模块在真包缺失时装一个基于 `xml.dom.minidom` 的替身，调用点完全不变；解析只是没有硬化，对「本工具链自己写的文档」这是可接受的取舍。

`_bootstrap.py` 是把这一切接起来的唯一入口，**每个新脚本的第一行都是 `import _bootstrap`**——一次 import 同时覆盖 `sys.path` 自定位、minidom 兜底、UTF-8 stdio。`encoding.py` 是最后一件事的具体实现；只想拿 stdio 修复、不要其余部分时才直接 import 它。

## references/：细节在这里，别塞进本文件

本文件写的是**决定**。实测出来的细节——做法、坑表、schema 原文——在 `references/`，需要时才读。

| 文件 | 什么时候读 |
|---|---|
| `references/formula-guide.md` | 写公式链生产代码、调 `_add_mixed`、排查 Word 报损坏、处理 `%`/`℃`/`～`/绝对值/复合单位等源文本坑 |
| `references/table-geometry.md` | 表格列宽不生效、要决定 `set` 还是 `sync`、`w:zoom` 缺 `w:percent` |
| `references/font-scheme.md` | 中英文混排字体规范、`w:rPr` 顺序与去重、`--heading-threshold` 语义 |
| `references/merge-and-export.md` | 合并后版心/编号/图片出问题、Word COM 导出细节、PIL 多页 PDF 写法 |
| `references/revisions-and-comments.md` | 批注与修订的 OOXML 表示与命令行语义 |
| `references/toc.md` | 写静态目录：两遍法（先 `page-probe` 量页码、再写条目）、制表位与点引线、全套坑表 |

## 目录结构

```
scripts/     平铺，按裸模块名 import —— 26 个 .py + 1 个 .ps1（25 个 CLI，1 个纯库）
references/  较长的做法笔记，按需加载 —— 6 篇
tools/       Word COM 验收辅助；--page-probe 在这里，即 word_accept.py
tests/       pytest 套件 —— 47 个 .py（46 个 test_*.py + _omml.py 辅助）
```

## 退出码

| 码 | 含义 |
|---|---|
| `0` | 成功 |
| `1` | 预期内的失败 —— marker 没命中、未知或重名的书签、文档里没有你要的东西、运行期异常 |
| `2` | 用法错误 —— argparse、spec 里出现未知键或未知 block type、`--widths` 个数多于列数、`--pages` 个数与目录条数不等、互斥 flag 同传、`--index` / `--insert-at` 越界或负数、输入路径不存在/是个目录/不是 Word 包 |
| `3` | **仅 `docx_mathify`**（以及把它的码透传出来的 `docx_template apply --fix-formulas`）：文档已写出，但至少一条公式没转成 |

用法错误是**拒收，不是修补**：判定发生在任何读写之前，输入文件逐字节不变，消息是 stderr 上的一行 `error: …`，stdout 保持空。

第 `2` 行里那条「输入路径」只约束**会打开文档的 CLI**，不是「仓里每个脚本都这样」：`docx_validate` 是刻意的例外——「这个包坏了」是它的**答案**，不是用错工具，所以它往 stdout 打 `{"ok": false, …}` 并 exit `1`，读它的 JSON 而不是它的退出码。其余那些以 `1` 报出来的失败（marker 没命中、书签不存在）都发生在**路径已经被接受之后**，那正是第 `2` 行划的界线。

`docx_mathify` 返回 `3` 是**故意的**：输出文件有效、绝大多数公式已转，但至少一条公式还是线性文本。**交付前必须查退出码**，只看文件存不存在会把半个公式发出去。「缺 XSL」并不必然退 `3`——文档里全是可降级片段（文字 run / 文字角标）时压根不碰 XSL，仍退 `0`。

`docx_read` 是 stdout 这条规矩最锋利的例子，因为它**每一种**成功模式都往 stdout 写 JSON：文件不存在、给的是目录、不是 ZIP、是 ZIP 但不是 Word 包，四种情况一律 exit `2`、stderr 一行、**stdout 一个字节都没有**。把失败当成 JSON 从 stdout 回传，会让每个 `json.loads(stdout)` 调用方把失败读成成功。「文档是真的、但里头有问题」归 `docx_validate` 管，那种走 exit `1`。

## 约定（Conventions）

- 每个脚本自己把所在目录塞进 `sys.path`，所以任意工作目录都能跑 —— `python scripts/x.py` 与 `python x.py` 都行。
- 机器可读的脚本往 `stdout` 写 JSON，给人看的写散文。诊断信息与错误一律走 `stderr`。
- 退出码反映结果；解析输出之前先查退出码。
- 文件读写一律显式 UTF-8；Windows 上 stdio 由 `_bootstrap.py` 切成 UTF-8，不要再手写 reconfigure。
- `requirements.txt` **故意纯 ASCII**：pip 用系统 locale 读它，一个 em dash 就能让 CJK Windows 装不上。

## 硬规矩

- 报告/正式文档的公式一律走 LaTeX→OMML；只有 `H₂O` 级极简上下标才用 `vertAlign` run。
- 所有脚本都能 `python -X utf8` 直接跑；Windows 上 stdout 由 `_bootstrap` 切 UTF-8，不要再手写 reconfigure。
- 就地改写前先备份。`docx_mathify --inplace` 会写 `<名>.bak.docx`，`docx_template.py apply --to` 不带 `-o` 时写 `<文件>.bak`；`fix_fonts.py`、`fix_zoom.py`、`merge_docs.py` 不留副本，原件重要就自己先复制。
- 需要 node 的只有 `docx_template.py apply -o`（docx-js 骨架）；断链时用 `apply --to`。