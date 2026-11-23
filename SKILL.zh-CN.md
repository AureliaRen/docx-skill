---
name: docx
description: "纯离线 Word .docx 工具集：创建/编辑/读取/校验/批注/修订/模板套用，外加 LaTeX→OMML 原生公式链与表格几何三处一致。当用户要生成或改 Word 文档、批量替换文档里的公式、修表格列宽/字体、合并 docx 或导出 PDF 时使用。另有自动编号、题注与交叉引用、静态目录、分节页码四组能力，供需要交付即正确（打开就看得见、F9 前后一致）的正式文档使用。"
license: MIT
---

# docx — 离线 docx 处理

`scripts/` 全平铺，**28 个 `.py` + 1 个 `.ps1` = 29 个文件**。**27 个是 CLI**（各自带 `if __name__ == "__main__"`
守卫，能直接从命令行跑），**`_bootstrap.py` 是唯一的纯库**；这 27 个里 **20 个有 argparse 表面**（`--help` 打印 `usage: …`），另 7 个有 `__main__` 但没有 argparse——6 个只打印一行「是库模块，import 它」自述（`docx_pack` / `docx_utils` / `encoding` / `linear_to_latex` / `math_omml` / `minidom_compat`），`fix_phasor_dot.py` 跑的是自检。另有 `tools/word_accept.py` 做文档级外挂验收（Word COM 实测页码，`--page-probe` 只属于它）。一律 `python -X utf8 scripts/<名>.py ...`，参数以各文件 `main()` 为准。

输出约定按脚本而异，这是刻意的：

- **stdout 给 JSON** —— `docx_create`、`docx_read`、`docx_edit`、`docx_validate`、`docx_comments`、`docx_revisions`、`docx_section`、`docx_numbering`、`docx_crossref`、`docx_toc`。这些拿去解析。
- **人类可读文本** —— `docx_mathify`、`table_width`、`insert_formula`、`docx_template`（`fill` 除外，它往 stdout 写 JSON）、`fix_fonts`、`fix_zoom`、`merge_docs`、`pdf_export`、`docx_pack`（裸跑会告诉你该怎么 import）、以及 `docx_common --api`（函数名 + 签名清单）。这些是给你看的进度。
- **两者按需给** —— `docx_preview` 默认打人类可读的行，加 `--json` 打一个 JSON 对象；同一次运行两种格式不会混。

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
| 看 run 级格式（写单元格**之前**先读这个） | `python scripts/docx_read.py in.docx --runs` |
| 查找替换 | `python scripts/docx_edit.py replace in.docx --find A --replace B -o out.docx` |
| 改表格单元格 | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X -o out.docx` |
| 改单元格并**丢弃**其全部格式 | `python scripts/docx_edit.py set-cell in.docx --table 0 --row 1 --col 2 --text X --replace-cell -o out.docx` |
| 插段落 | `python scripts/docx_edit.py insert in.docx --index 3 --text "..." -o out.docx` |
| 删段落 | `python scripts/docx_edit.py delete in.docx --index 3 -o out.docx` |
| 套段落样式 | `python scripts/docx_edit.py style in.docx --index 0 --style "Heading 1" -o out.docx` |
| 合并格式相同的相邻 run | `python scripts/docx_edit.py normalize in.docx -o out.docx` |
| 给锚点文字设高亮 / 颜色 / 粗体 / 斜体 | `python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta --bold -o out.docx` |
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
| 渲染 + 版面自检（PDF、每页 PNG、半空页报告） | `python scripts/docx_preview.py in.docx --pdf out.pdf --pages png/ --lint --json` |

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

**`pagenums --template` 有五种形态**，中文交付用 `zh`：`第 <PAGE> 页　共 <NUMPAGES> 页`。`页` 与 `共` 之间那个分隔符是**全角空格 U+3000**，不是普通空格——半角那道缝在印刷尺寸上读起来像排版事故。加 `--total sectionpages` 就把第二个域换成 `SECTIONPAGES`，于是「共 Y 页」数的是本节而不是整篇。`zh` 会忽略 `--prefix` / `--suffix`（它自己拼文本），但**不会静默吞掉**：返回的报告里多一条 `"ignored": ["prefix"]` 自报。

总页数是**域，不是字面文字**，这决定了怎么验它。`第 1 页　共 <NUMPAGES> 页` 落到 `footer*.xml` 里时，总数那一段的缓存值是 `1`——因为包还在写的时候，文档一共几页根本不知道；Word 与 LibreOffice 打开文档时才重算这个域。于是五页的文档，PDF 上写 `共 5 页`，XML 里仍是 `1`。**验页脚要看渲染结果，不要看原始 XML**：拿「XML 里应该是 5 页」去查，会在一个完全正确的文档上报错。

`--template custom` 拿 `--prefix` / `--suffix`，并把里面的 **`{page}`** 与 **`{total}`** 按出现顺序展开成 PAGE 与 total 域，所以 `--prefix "{page}/" --suffix " of {total}"` 得到 `1/ of 12`。两个串里都没有占位符时，`custom` 的行为与从前逐字节一致。同一套键在 spec 里同样成立：顶层 `footer_page_numbers`，或者 `section_break` 里的那一个，都接受 `"template": "zh"`，也接受带占位符的 `custom`。

**`docx_template.py apply` 带两条独立的公式通道、两套各自的默认值，命令行上互斥：**

| flag | 通道 | 默认 |
|---|---|---|
| `--no-fix-formulas` | marker 通道 —— 成对 `_xx_` / `^xx^` → `vertAlign` run | **开** |
| `--fix-formulas` | OMML 通道 —— LaTeX / 线性公式 → 原生 Office Math（内部跑 `docx_mathify`） | 关 |

这两个名字看着像一对 `+`/`-`，实际写的是两个不同的 dest（`fix_formulas` / `omml_formulas`），所以早先同传会一边关掉 marker 通道、一边把裸 `_xx_` 标记交给 OMML 转换器 —— 转换随之失败，而 CLI 仍退 `0`。现在是 argparse 冲突，exit `2`。不给 `-o` 时 `apply --to` 就地改写目标、但先写一份 `<文件>.bak`；给了 `-o`，`--to` 那份原文件一个字节都不动。

**`--fix-formulas` 会把 `docx_mathify` 的退出码原样往上报，包括 `3`。** OMML 通道是就地跑一遍 mathify，而那边的 `3` 意思是「文档已写出，但至少一条公式还是线性文本」—— 唯一一个「有产物 + 非零码」的出口。这里若仍旧报 `0`，就等于把半成品当完成品交给调用方。所以现在 `apply` 返回的就是 mathify 返回的那个值：`1` 和 `2` 原样透传，`3` 也透传为 `3`，只有 `0` 才归零。`--no-fix-formulas` 压根不调 mathify，行为与从前一致。

## 创建文档

`docx_create.py` 读一份 JSON spec。顶层键：`page`、`styles`、`header`、`footer`、`footer_page_numbers`、`blocks`。每个 block 有一个 `type`：

`footer_page_numbers`（顶层或 `section_break` 里）接受 `fmt`、`start`、`align`、`template`、`total`、`prefix`、`suffix`、`page_type`，以及 `font`、`size_pt`、`bold`——这三个作用于承载页码的 run：`font` 与 `size_pt` 文字 run 和域 run 都吃，`bold` 对承载页码数字的域 run 只在真值时生效（精确口径见下表）。它们早先一设就抛 `ValueError`，现在可用；中英混排的页脚里 `font` 会把四个 `w:rFonts` 槽（`ascii` / `hAnsi` / `eastAsia` / `cs`）一起绑死，`第 X 页　共 Y 页` 这种中文前缀才不会退回主题的东亚字体、让数字单独用你指定的那个。

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
| `crossref` | `target`（书签名），可选 `prefix`、`suffix`、`format`（`default`/`full`/`above-below`/`page`/`number`/`paragraph-number`/`relative-context`/`full-context`）、`cached`、`own_paragraph`、`superscript` / `vert_align`、`hyperlink`、`static`、`no_mergeformat` |
| `table` | `header`、`rows`，可选 `style`（默认 `Table Grid`）、`header_bold`、`keep_together` |
| `image` | `path`，以及 `width_mm` **或** `width_frac`（二者互斥） |
| `page_break` | — |
| `section_break` | `start_type`，可选 `page`、`header` / `footer`（及其 `_align` / `_font` / `_size_pt` / `_bold`）、`footer_page_numbers`（`fmt` / `start` / `align` / `template` / `total` / `prefix` / `suffix` / `page_type`，以及给页码 run 用的 `font` / `size_pt` / `bold`——见上面那张 `bold` 表：`false` 管得到文字 run，管不到域 run）、`different_first_page`、`cols`、`no_footer`、`reuse_last_paragraph`、`apply_to` |
| `toc` | 可选 `instr` —— 默认 `TOC \o "1-3" \h \z \u`；要图目录就传 `TOC \h \c "图"`。**这是一个域**：Word 必须更新（F9）之后才有条目 |

**spec 在写出任何一段之前先被校验；未知键与未知 block type 都是错误，不是 no-op。** 早先 `{"blockz": [...]}` 会造出一份合法但空白的文档、打印 `"blocks": 0`、退出 `0`，作者根本看不出内容没进去。现在一律 exit `2`，一行报错点名出问题的键，拼写接近时还给提示：

```
spec: unknown key "blockz" (did you mean "blocks"?); valid keys here: blocks, footer, footer_page_numbers, header, page, styles
blocks[0]: unknown block type 'paragraf' (did you mean "paragraph"?); choose one of: bullet_list, caption, crossref, heading, image, numbered_list, page_break, paragraph, section_break, table, table_caption, toc
```

exit `2` 的承诺是**什么都没写**，所以改完 spec 直接重试，不用猜是不是留了半份文档。值也不做钳制：超范围的值会连同它本该被钳进去的区间一起报出来。

**spec 的键名必须是目标命令真正读的那个名字，这一条不是白来的。** 校验器只证明键按**本表**的拼法写对了；随后 builder 把值经一个 namespace 对象交给 `cmd_ref` / `cmd_caption`，而它们用 `getattr(args, "name", 默认值)` 取值——**这个写法分不出「没有这个属性」和「属性是 false」**。所以名字对不上就是**静默 no-op**：修复前实测，`{"type": "crossref", "target": "g1", "hyperlink": false}` 依然产出一份含 `<w:hyperlink>` 的文档，`exit 0` 且 `{"ok": true}`。spec 里叫 `hyperlink`（正向，默认 true），`cmd_ref` 读的是 `no_hyperlink`（反向），所以 builder 必须**翻译而不是转发**。`tests/test_create_crossref_args.py` 从 `docx_crossref.py` 里把 `getattr` 的名字解析出来做全量对照，builder 少传一个就红。

`styles[]` 条目接受 `font`、`font_east_asia`、`size_pt`、`bold`、`italic`、`color`、`base`、`outline_level`——最后一个会写 `<w:pPr><w:outlineLvl>`，Word 才把该样式当标题。缺了它，模板自带的 `卷标题` / `ChapterTitle` 在 Word 眼里就是**正文**，长得再像也没用：导航窗格是空的，`TOC \o` 一条也收不到。

```json
{"styles": [
  {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
   "bold": true, "outline_level": 1}
]}
```

`outline_level` 是 1 基（`1` → `w:outlineLvl val="0"`）。`0` 表示「显式正文」（`val="9"`），给那种永远不该被收集的样式用。

**`styles[]` 条目点名一个已存在的样式时，是就地修改，不是报错，而且只动你给出的键。** `{"name": "Heading 1", "size_pt": 16}` 改的是 Word 内置 `Heading 1` 的字号，它的字体、粗细、颜色、`outlineLvl` 一律原样留着；同一个条目再 apply 一次，`<w:style>` 仍然只有一个。没给出来的键永远不会被重置，所以不会有东西悄悄退回原值。想把颜色改回去就显式写 `color`——另外注意：本工具生成的文档里 `Heading 1`–`Heading 9`、`Title`、`Subtitle` **默认就是黑色**（见 [标题默认是黑的](#标题默认是黑的)），所以你没染过色的标题本来就是 `000000`，不是主题蓝。

`heading.level` 是 `0`–`9`，不是 `1`–`9`：python-docx 的 `add_heading` 把 `0` 当 Word 的 **Title** 样式、`1`–`9` 当 Heading 1–9。`10` 是 spec 错误，不会被钳。

### 标题默认是黑的

本工具生成的每一份文档，在套用任何 `styles[]` 条目**之前**都会先过一遍 `Heading 1`–`Heading 9`、`Title` 与 `Subtitle`：摘掉主题色绑定（`w:themeColor` / `w:themeShade` / `w:themeTint`），写上 `w:val="000000"`；四个字符变体（`Heading1Char` … `TitleChar`、`SubtitleChar`）同样处理；`Title` 的段落下边框换成黑色（`w:pBdr` 上写 `w:color="000000"`，`w:sz` 与 `w:space` 不动）；`Subtitle` 去掉斜体（`w:i` 与 `w:iCs`）。理由是：标题不该因为某台机器的主题色不同而变色，生成的文档也不该在没人要蓝的时候自带主题蓝。

这是**默认值，不是锁**。因为中和在 `styles[]` 之前跑，所以给了 `color` 的条目照样覆盖——就地 modify 打开的那道门在这里一样有效。`Heading 9` 的斜体保留；去掉斜体的只有 `Subtitle`。

### 让一张表不跨页

`{"type": "table", …, "keep_together": true}` 要求 Word 把整张表压在一页。Word 没有「表级 keep-together」这种属性，所以请求分两层写：**每一行**的 `w:trPr/w:cantSplit`（一行不许被分页切两半），加上除末行外每一行每一段的 `w:keepNext`（每行把下一行拽住）。默认 `false`，不写这个键时上述元素一个都不会生成。

**`cantSplit` 只禁「拆行」，一整页都放不下的表仍然会跨页。** `keep_together` 是请求不是保证：它救的是那种「本来只差两行、结果被切到下一页」的 8 行表，对一页根本装不下的 60 行表无能为力。

`table_caption` 紧跟一个 `table` 现在**默认联动**，不用任何键：caption 段拿到 `keep_with_next`、表首行拿到 `cantSplit`，于是单行题注不会被孤立在上一页页底。中间插任何东西——段落、分页符、另一张表——都会断掉联动，因为「紧跟」是按 block 列表量的，不是猜的。这个联动会改变在它之前写的 spec 的排版结果，这正是目的；真需要旧行为，就在 caption 与表之间放一个段落。

### 按版心比例给图片定宽

`image` block 用 `path` 加上 `width_mm`（绝对宽）**或** `width_frac`（当前节正文宽的比例，`0 < f ≤ 1`）之一，绝不同时给——两个都给是 exit `2`、零写盘。`width_frac` 按 `content_width_mm(section)` 折算，所以它跟着纸张与页边距走，而不是写死一个数：A4 且左右各 26 mm 时正文宽 158 mm，于是 `1.0` 是 158.0 mm、`0.5` 是 79.0 mm。越界的值连同区间一起拒收、不做钳制；`0.0` 也拒收，而不是读成「不定宽」——调用方写下的数字，永远不会被悄悄换成别的意思。

`path` 按**当前工作目录**解析，**不是**按 spec 文件所在目录；路径不存在会被当成输入错误当场拒收——exit `2`、消息里带 block 序号、零写盘，而不是留到写文档的过程中炸出一个 `FileNotFoundError` traceback。

### 让标题从新的一页开始

任何 block 的 `page_break_before: true` 都会往它的 `<w:pPr>` 里写 `<w:pageBreakBefore/>`，这正是 Word 自己在*段落 → 换行和分页 → 段前分页*里写的元素。要用就用这个键，挂在标题上。

**别拿 `{"type": "page_break"}` 去顶替它。** 分页 block 是一个只装分页 run 的空段落，它待在哪取决于 block 顺序：标题前多一个、或者连着两个，就多出一页空白或者把内容顶过一页——而正文中间凭空一页空白，正是 `docx_preview.py --lint` 会报成 `blank_pages` 的那种毛病。`page_break_before` 把分页绑在它所属的段落上，那个段落挪到哪，分页就跟到哪。

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

### `ref --format` 的八个值，以及各自写的是哪个域

`--format` 决定这条引用落成哪个 Word 域。合法值一共八个：前五个管日常，后三个是文档真需要时才出现的。

| `--format` | 写出的域 | 读起来是 |
|---|---|---|
| `number` | 裸 `REF <bookmark>` 编号 | `3` |
| `default` | `\h` —— 指向书签的超链接 | `3`，可点 |
| `full` | 整段题注文字（要求书签跨住那一段） | `图 3 一张…` |
| `above-below` | `\p` —— 位置相对词 | `上图` / `下图` |
| `page` | `PAGEREF` —— 书签落在第几页 | `7` |
| `paragraph-number` | `\n` —— 段号 | `[1]`，就是参考文献列表上那个编号 |
| `relative-context` | `\r` —— 相对最近一个编号项的编号 | 相对所属标题 |
| `full-context` | `\w` —— 书签的完整上下文 | 编号加上它所在的上下文 |

四个开关的差别是「一个数字」和「读者能据此行动的数字」之间的差别：`\p` 回答**在哪一侧**，`PAGEREF` 回答**在第几页**，`\n` 回答**第几个编号项**，`\r` 回答**相对于谁**。**按句子在问什么挑域**，别按哪个输出看着顺眼挑。

`--format full` 要求书签跨住那段。`--format page` 不能与 `--static` 同用（见下）：页码在分页发生之前不存在。

### 上标引用：域还是静态，各自付什么代价

「图¹」这类上标引注，是两种设计差别最明显的地方，而且**不是外观问题**。`--superscript` 写的是 `<w:vertAlign w:val="superscript">`，它能不能活到 PDF 里，取决于你选了下面哪种形状。

| | `ref`（默认，域） | `ref --static` |
|---|---|---|
| 标记 | `<w:hyperlink>` 包着 5 个 run 的 `REF` 域 | `<w:hyperlink>` 包着**一个** `<w:r>`，编号就是 `<w:t>` 文本 |
| 编号跟着题注走 | 会——F9 / 打印自动重算 | **不会**——写的时候就冻住 |
| 熬得过 `Fields.Update()` | 熬得过，**但必须有 `\* MERGEFORMAT`** | 无条件熬得过：没有域可重绘 |
| 要不要先按 F9 | 不用（缓存结果已写） | 永远不用 |
| 在它上面加一张图 | 自动重编号 | 得重跑命令 |

咬人的是第三行。Word 16.0 实测，整篇 `Fields.Update()` 后导 PDF，用 PyMuPDF 读回 span：

```
域 + \* MERGEFORMAT     '1'  sz 7.0  y 197.18     上抬 2.52 pt、半字号
域 + --no-mergeformat   '1'  sz 11.0 y 199.70     上标没了
--static                '1'  sz 7.0  y 197.18     与第一行一致
正文                    sz 11.0 y 199.70          基线，供对照
```

所以 `MERGEFORMAT` **不是装饰**：没有它，Word 会从书签目标重新渲染域结果，下一次更新就把上抬的效果抹掉——**而导 PDF 正好就是一次更新**。它默认就写；`--no-mergeformat` 只用来复现这个故障。

`--static` 用在你要让外观**结构性**成立、而不是靠一个开关撑着的时候——比如流水线上还有别的东西会重建域，或者文档必须在不更新域的阅读器里也是对的。代价要清楚：编号从此不跟题注走，增删图表就得重跑。**文档还在改的时候用域形式（默认）**；`--static` 是收尾那一遍用的。

`--static` 会拒绝 `--format page` 与 `--format above-below`（exit `2`）：页码依赖分页，建文档时页面还没排出来；上方/下方由 Word 逐次判定。把这两个冻成字面量，等于写一个「看起来像对的错数」。

两种模式都会在报告里说明自己用的是哪种：`"mode": "field"` 配 `"needs_field_update": true`，或 `"mode": "static"` 配 `"needs_field_update": false`。构建脚本可以直接断言拿到的是想要的那种，而不是从标记反推。

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

### 什么能作为证据（以及分数符号）

一个 span 要自己挣进入场资格。默认要有关系符（`=`、`≈`、`≤`、`±`）或强算子（`√`、`∑`、`∫`、`∂`、`∇`）。有两样豁免，因为它们是**作者自己声明的**，不是我们猜的：

- 显式角标组 —— `S_{甲}`、`m^{2}`；
- **vulgar fraction 分数符号** —— `½ ¼ ⅓ ⅕ ⅛ ⅐ ↉` 及其余 11 个。

第二条值得单独说，因为这个字符是**隐患**而不是错误。U+00BD 是单个码位，字体把它渲染成小号 `1` + 斜杠 + 下沉 `2`：在一种字体里**看着差不多**、换一种字体就明显不对，于是没有任何一处会报错，问题一直活到打印。19 个 Unicode 分数码位现在全部翻成 `\frac{n}{d}`，而且**单独一个就足以开一个 span** —— `½+1` 没有 `=` 也照样转，`第½章` 里那个孤立的 `½` 也转，两侧的汉字留作文本。

分数符号同时是单位文法的**数值锚点**，这正是两种含义的分界：

```
½ m/s     → \frac{1}{2}\,\text{m/s}     测量值：「每秒半米」
0.5 m/s   → 0.5\,\text{m/s}             同一个量，同一种排版
v=m/s     → v=\frac{m}{s}               公式：m 除以 s
m/s       → \,\text{m/s}                裸的单位串
```

这一列背后的判据**不是**「两边都像单位」—— `m/s` 与 `v=m/s` 在这条测试上完全一样。判据是锚点：左边有数值（`0.5`、`10⁻⁷`、`½`、右括号或幂次）时，单位串是**测量值**。所以 `½ m/s` 与 `0.5 m/s` 走单位路径，`v=m/s` 不走、仍作分式。

看中间产物时有一条要知道：间距命令（`\,` `\;` `\:` `\!` `\ `）是**一个 token**，不是「反斜杠 + 标点」。必须是 —— 否则分式的操作数扫描会把那个反斜杠并进分母，花括号提前收口：`1/2 m/s` 会出成 `\frac{1}{2\},\text{m/s}`，分母渲染成字面的 `2},`。这是本轮之前一直存在的 bug；触发形态是「分式后紧跟单位串」，而那正是物理讲义的大半篇幅。


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

## 改表格单元格而不毁掉它的格式

`set-cell` 塞进单元格的是**一段平字符串**。模板里的格很少是平的：一个格里常有
好几个 run，而要命的那点格式往往只在其中一个 run 上。像这样一行

```
v₀ = 12.5 m/s      ← 三个 run："v" 常规、"0" 下标、"= 12.5 m/s" 加粗
```

绝不能用 `cell.text = ...` 覆盖。那个 setter 的文档原话是「Any existing content
or revisions are replaced」：它清空格子、加一个段落、加一个 run，然后**把所有
run 级属性全丢掉**——段落对齐、字号、字体、`vertAlign`，一概不留。更糟的是它
干得悄无声息：`exit 0`、零警告，等你用 Word 打开才发现整行变成了下标。

所以 `set-cell` 默认走 `mode: "preserve"`，并且**如实报告它动了什么**：

| 情形 | 行为 | `mode` |
|---|---|---|
| 格里只有 1 个可写 run | 替换该 run 的文字，`rPr` 原样不动 | `preserve` |
| 格里有多个 run | 文字写进**第一个**可写 run，其余 run 被移除；该 run 若带 `vertAlign` 会被剥掉，免得新文字继承一个野下标 | `preserve` + `warnings` |
| 传 `--replace-cell` | 旧行为，一模一样：清空、一个段落、一个 run、零格式 | `replace-cell` |

**不可写的 run** 指域（`w:fldChar` / `w:instrText`）、`drawing`、图片——这些一律
保留而不覆盖，因为换掉产生域结果的那个 run，域就废了。

写之前先把格子读清楚：

```bash
python scripts/docx_read.py in.docx --runs
```

它会按段落打印 `mixed_format` 标志，以及每个 run 的 `text` / `vert_align` /
`bold` / `italic` / `size_pt` / `font` / `is_field` / `is_drawing`，并带一个形如
`table[0].r0.c0` 的 `path` 供你定位。`mixed_format: true` 的意思是「一段平字符串
表达不了这个格子」——要么明知后果地传 `--replace-cell`，要么用 `python-docx`
深拷贝要保留的那个 run 再重写（姿势见 `references/editing-traps.md`）。

凡这次没能保住的，都会在 stderr 打一行 `warning:`，同时进 JSON 的 `warnings`
数组。**`exit 0` 不等于格子出来跟原来一样。**

## 字符级格式：给锚点文字设高亮 / 颜色 / 粗体 / 斜体

`set-cell` 和 `replace` 写的是**文字**。`char-format` 只改**已有的字**长什么样，
而且按内容定位，不按下标：

```bash
python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta
python scripts/docx_edit.py char-format in.docx --find "重点" --color 00a0ff --bold -o out.docx
python scripts/docx_edit.py char-format in.docx --find "关键词" --bold --italic -o out.docx
```

**选择器是锚点字符串，而且命中几处就改几处。** `--find` 逐段比对（正文、表格、
嵌套表格、页眉页脚），每一处都算 —— 两段各含一个 `重点` 时实测报告
`{"occurrences": 2, "runs_changed": 2, "paragraphs_touched": 2,
"paragraphs_searched": 8}`。这里刻意**没有** `--index`：`docx_read --runs` 的段落
编号横跨正文 + 单元格 + 文本框，而 `doc.paragraphs` 只数正文段落，同一个工具对里
「index 7」会指两个不同的段落。文字也是更耐久的定位器 —— 任何一次编辑都会让后面
的位置整体平移，而平移后的位置照样「能用」，只是格式落到了别的行上。

**锚点只占某个 run 的一部分时，会把这个 run 切开**，前后文字保住自己原有的格式。
实测：`本段包含重点两个字，其余不变。` 上执行 `--find 重点` 之后，`--runs` 报出三个
run —— `本段包含` / `重点` / `两个字，其余不变。`，只有中间那个带上了新属性。切不动
的 run（里面还有制表符、换行、域或图片）会被**拒收**而不是整块改掉：「比你要求的
多高亮了几个字」和 bug 在调用方眼里分不出来。

### `--highlight` 的合法值，以及 python-docx 叫不出名字的那个

合法 `w:highlight` 值共 17 个（`ST_HighlightColor`）：

```
black blue cyan darkBlue darkCyan darkGray darkGreen darkMagenta darkRed
darkYellow green lightGray magenta none red white yellow
```

**拼法自由，值不自由。** `darkMagenta` / `DARK_MAGENTA` / `dark_magenta` 以及
python-docx 的成员名 `VIOLET` 都会归一成 `darkMagenta` —— 实测 `--highlight
violet` 退出 `0`，落盘是 `w:highlight w:val="darkMagenta"`。走别名路线是刻意
设计的：`WD_COLOR_INDEX` 有 18 个成员，其中只有 16 个在这里合法（`AUTO` 写的是
`default`、`INHERITED` 什么都不写），而且它**没有 `darkMagenta`** —— 恰恰是这个
功能被要求支持的那个颜色；`VIOLET` 用一个人人都不叫的名字写它。所以那张枚举只
被当作**别名来源**读取 `xml_value`，从不拿来做闸门；闸门是那 17 个值。

报告会把**归一后**的值同时放进 `applied` 和顶层，确认「到底落进去什么」不必去翻
XML：

```json
{"find": "重点", "occurrences": 2, "runs_changed": 2, "paragraphs_touched": 2,
 "paragraphs_searched": 8, "applied": {"highlight": "darkMagenta", "color": "00A0FF", "bold": true},
 "highlight": "darkMagenta", "color": "00A0FF", "bold": true, "ok": true, "output": "out.docx"}
```

`--color` 收六位十六进制（带不带 `#` 都行，如 `FF0000`、`#00a0ff`）或 `auto`；
写盘时转大写（`#00a0ff` 落成 `00A0FF`），并且会替换掉该 run 上的主题色 —— 在
`CT_Color` 里主题属性在渲染时压过 RGB 孪生值。

### `--highlight none` 与 `--highlight inherit` 不是一回事

在一个干净的 run 上两者看着都是「没有高亮」，所以这个区别必须写清楚。实测在一个
**本来就带** `w:highlight w:val="yellow"` 的 run 上：

| 旗标 | 调用后的 `w:rPr` | 含义 |
|---|---|---|
| `--highlight none` | `<w:rPr/>` —— 元素被**删掉** | 明确取消高亮；样式层的高亮会被压住 |
| `--highlight inherit` | `<w:rPr><w:highlight w:val="yellow"/></w:rPr>` —— 原样不动 | 什么都不写；样式怎么说就怎么算 |

`inherit` 就是「这个旗标别碰」：它的存在是为了让一次调用可以只设 `--color`（或
`--bold`），而不顺手清掉别人放的高亮。`none` 是**主动删除**。两者都是
`--highlight` 的合法*取值*，而且都不写假值。

粗体斜体是同一套三态逻辑，`w:val="0"` 表示显式关：`--bold` 写 `<w:b/>`，
`--no-bold` 写 `<w:b w:val="0"/>`（实测 `<w:rPr><w:b w:val="0"/><w:highlight
w:val="yellow"/></w:rPr>`），两个都不给就什么都不写。只写拉丁孪生 `w:b` / `w:i`，
从不写 `w:bCs` —— 一个和 `w:b` 打架的 `w:bCs` 本身就是另一种惊吓。

### 拒收：`exit 2`、stderr 一行、一个字节都不写

下面每一条都发生在 `doc.save()` **之前**，所以输入文件逐字节不变、stdout 保持空：

| 情形 | stderr（实测原文） |
|---|---|
| 值不在那 17 个里 | `error: --highlight 'purple' is not a legal w:highlight value; legal: black, blue, cyan, … (any spelling, e.g. darkMagenta / DARK_MAGENTA / dark_magenta)` |
| 没有任何 run 带这个锚点 | `error: char-format --find '不存在的字': no run in the document carries this text. Searched 8 paragraph(s) across body, tables, headers and footers (text boxes and tracked-change content are out of scope, as for every other run-level writer).` |
| 一个格式旗标都没给 | `error: char-format needs at least one format to set: --highlight, --color, --bold/--no-bold, --italic/--no-italic` |
| `--color` 不是十六进制 | `error: --color 'red' is not a colour: expected 6 hex digits (FF0000 or #00a0ff) or 'auto'` |

拼错的颜色从前是**原样写进去**的，结果是文档能校验、能打开、**一点高亮都不显示**，
而且没有任何东西解释为什么。锚点一处都没命中同样拒收，理由和 `replace` 那条一样
（见[第 4 条：改文本不要用 shell 的字符串替换](#4-改文本不要用-shell-的字符串替换)）：
静默的零看起来就像成功。

### 验证靠读回

设置不是证据，读回才是。`docx_read --runs` 会逐 run 报 `highlight` 与 `color`，
正好是这条命令写的那两个：

```bash
python scripts/docx_edit.py char-format in.docx --find "重点" --highlight darkMagenta -o out.docx
python scripts/docx_read.py out.docx --runs
```

实测切开后的读回：带 `重点` 的那个 run 报
`{"text": "重点", "highlight": "darkMagenta", "color": "00A0FF", "bold": true}`，
左右两个邻居三项全是 `null`。**邻居为 null 才是真正的检查** —— 它证明切 run 没有
把格式渗到周围文字上。

## 浮动文本框

文本框里的段落**不是** body 段落。它们住在 `w:txbxContent` 里，父链只可能是
`wps:txbx`（DrawingML `w:drawing` 内）或 `v:textbox`（VML `w:pict` 内），
**永远不是 `w:body` 的直接子元素**。于是 `doc.paragraphs` 不列它们、
`body.iterchildren()` 到不了它们、`Paragraph.text` 也只停在自己的 `w:r` 子元素上：
**任何沿 body 段落遍历的模式都会把这种文档读成空的**。真实文件实测：一份阅读
批注作业有 43 条批注全在浮动框里，`--text` / `--blocks` / `--runs` 三个模式
**各读到 0 条** —— 这就是「用 `python-docx` 扫一遍，判定 43 处全是色块、没有
批注文字」的来源。

四种读取模式现在都覆盖它们：

| 模式 | 框里的文字出现在哪 |
|---|---|
| `--text` | 新增独立键 `textboxes`（一个框一个字符串）；`body` 的语义不变 |
| `--blocks` | 紧随宿主段落之后的 `textbox` 块，带 `index` / `flavor` / `text` / `paragraphs` |
| `--runs` | 框自己的段落，每段带 `textbox: <n>`，`path` 形如 `body textbox[<n>]` |
| `--md` | 框里的文字作为引用块（`>`）输出 |

body 段落与表格单元格不受影响：无框文档的块序列与从前一致，`body` 的语义不变，
每段的 run 字段也一个不少 —— 新增的只有恒为空的 `textboxes` 列表和每条 `--runs`
记录上恒为 `null` 的 `textbox` 字段。最后这一点正是为什么「不变」不能读成
「逐字节相同」：`--runs` 多了一个键，所以它的输出是**超集**，不是同一批字节。

**有框**的文档里序号会平移，这没法避免：`textbox` 块插在宿主段落之后，于是插入点
之后的每一条都往下挪（实测：六个无框段落编号 0-5；四个框锚在第 1 段里，正文第 5
段的 `index` 从 `4` 变成 `12`——每个框既带来自己的段落记录，又带来一个 `textbox`
块）。跨一次编辑缓存 index 之前要知道这件事 —— 稳定的是块的 `type` / `text` /
`style` 这几个键，不是它的位置。

读输出时要知道两件事：

- **一个框一条**。Word / WPS 把框写成 `mc:AlternateContent`，里面 `mc:Choice`
  （DrawingML）与 `mc:Fallback`（VML）**各存一份同样的文本**。读者在有文本的
  Choice 里取，Choice 没文本时才退到 Fallback —— 所以 43 个框报 **43**，不是 86。
  自己写遍历时必须合并这一对，否则每个计数都会静默翻倍。
- **`index` 是框的序号**，`--blocks` 与 `--runs` 用同一个号码指同一个框。

分工：`docx_footer.py` 读页眉页脚里的框并量它的几何；`docx_read.py` 读正文里
框的**文字**。两边不重复实现。

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

## 工作流程：每次存盘之后的那几项检查

四条规矩，每一条都是真实交付事故换来的。有人给一份「阅读批注作业」做 Word 排版，成品要交。工具链就在手边，他一个检测器都没用，改走「转 PDF → 转 PNG → 肉眼看图 → 手动量 y 坐标」这一套，前后返工了**十几轮**。事后对成品跑一次 `--lint`，它当场说出了他肉眼追了十几轮的同一批问题：

```
lint: 2 issue(s), 5 page(s)
lint: sparse_pages=[3, 5] blank_pages=[]
  warning sparse-page page 3: fill_ratio 0.213 < 0.40 (content covers 21.3% of the page height; …)
  warning sparse-page page 5: fill_ratio 0.117 < 0.40 (content covers 11.7% of the page height; …)
```

十几轮手工丈量，一条命令说完。当前版本那行 `lint:` 还会带上 `content_width_mm`。下面按「单位时间收益」排序。

### 1. 每次存盘后跑 lint——最高优先级

```bash
python scripts/docx_preview.py FILE.docx --lint --text
```

只给 `--lint` 也会转一份临时 PDF（用完删掉），逐页量完再报 **8 个** issue code——常态报 7 个，加上 `--gaps` 再报 `tail-gap`。（`fill_ratio` 是逐页的数字，不算 code。）**先看输出的第二行**——`sparse_pages` 非空就是有半空页，那是没干完的活，不是一个可以放过的警告：

| JSON 键 | issue code（级别） | 含义与处置 |
|---|---|---|
| `sparse_pages` | `sparse-page`（warning） | `fill_ratio` 低于 `--min-fill`（默认 `0.40`）的页。**非空 = 必须处理**：挪块、缩图、删掉多余的换页 |
| `blank_pages` | `blank-page`（error） | 既无文字也无图片的页 |
| `table_split` | `table-split`（warning） | 表格跨到了下一页。启发式、允许误报——读成「这页值得看一眼」 |
| `orphan_last` | `orphan-last`（warning） | 该页正文末行只剩 1–2 个字符（`--max-orphan-chars`） |
| `image_over_width` | `image-over-width`（error） | 图比版心宽（对照 `content_width_mm`） |
| `issues` | `content-width-unknown`（warning） | 文档没报出节页宽，所以**图片宽度根本没比对**。这是**降级信号**，不是缺陷：此时 `image_over_width` 为空只说明「没量」，不说明「干净」 |
| `issues` | `section-margins-differ`（warning） | 多个节的边距不一致，而几何只取第 1 节，所以后面几节的 `head_gap_mm` / `tail_gap_mm` 是拿错的正文区边界量出来的 |
| `gap_pages`（需 `--gaps`） | `tail-gap`（warning） | `tail_gap_mm` 超过 `--max-tail-gap`（默认 `60`） |

**`exit 0` 不代表版面没问题。** lint 的结果是**输出**，不是失败：`main()` 打完报告就返回 `0`，与严重性无关。实测一份单页文件，末行只剩一个 `脚`：它打出 `warning orphan-last page 1 ends with a lone '脚' on its last line`，退出码仍是 `0`。所以要去 grep 报告里的 `sparse_pages`、`blank_pages`、`issues`，**绝不能拿退出码当判据**。这与 `docx_validate` 那一侧是同一个坑的镜像：那边 `1` 才是答案。

`fill_ratio` 给的是**比例**；`--gaps` 另给**毫米绝对值**。两个数回答的是不同问题——文字起始位置靠下时，某页可能只有 `0.213` 而页尾空白并不大。给人类看报告时引用毫米：`tail_gap_mm` 与 `head_gap_mm` 才是尺子量出来的那个数（「空了 8.3 厘米」），而且 `top_margin_mm + head_gap_mm` 可以拿去和渲染出的正文首行对账。

### 2. 断定「文档里没有 X」之前，必须换一个读者交叉验证

`python-docx` 只走一棵树，另一棵**静默跳过**。`paragraph.runs` 只覆盖**直属 run**：`w:txbxContent`（文本框内容）与 `mc:AlternateContent`（兼容包）不在那条路径上，会被无声丢掉、不报错。同一份批注文件实测：**43 条批注全在浮动框里**，`paragraph.runs` 扫出来是 **0 条**——于是被判定成「43 处全是色块、没有一处批注文字」，核心成果被当成主要缺陷，后续修改建议整个建在这个误判上。循环看着完全正确，结论是错的。

改用三种模式交叉看：

```bash
python scripts/docx_read.py FILE.docx --blocks   # 有序 body 流 + textbox 块
python scripts/docx_read.py FILE.docx --text     # 正文 + 独立 textboxes 列表
python scripts/docx_read.py FILE.docx --runs     # 每段 run，带 textbox: <n>
```

三种现在都能读到浮动框——`--text` 有独立键 `textboxes`，`--blocks` 出 `textbox` 块，`--runs` 给框内每段打 `textbox: <n>`。机制、`mc:Choice` / `mc:Fallback` 计数翻倍的坑、以及两边共用的框序号，都在[浮动文本框](#浮动文本框)那一节；自己写遍历之前先读它。

原则一句话：**「我扫了一遍没看到」不等于「文档里没有」——换工具再扫一遍。** 一个读者报 0，量的是那个读者，不是那个文件。

### 3. 存盘后跑 validate

```bash
python scripts/docx_validate.py FILE.docx
```

`docx_validate` 回答 lint 答不了的那个结构问题：这个包本身是否健全。它能抓空 `<w:drawing/>`（Word 直接拒开）、内容部件的**非法嵌套**、悬空关系（`dangling-rel`）、`r:embed` / `r:id` 找不到关系（`unresolved-reference`）、引用了不存在的样式（`missing-style`）、run 属性里 Word 不收的值（`invalid-run-property` / `nonstandard-run-property`，见下），以及页脚那几项：浮动页脚框窄到装不下那一行字（`footer-textbox-overflow`——实测一个 4.56 pt 的框被要求装 121 pt 的文字）、页脚没有 PAGE 域（`footer-no-page-field`）、页脚部件读不出来。

### `invalid-run-property` 与 `nonstandard-run-property`

`<w:rPr>` 里那些 `w:val` 取自封闭值表的子元素（`w:highlight`、`w:vertAlign`、`w:effect`、`w:u`、`w:color`）按 **Word 实际怎么处理**来判，不按 schema 判。这一刀切出两个 code，分界线只有一个：**文件能不能打开**。

| code | 严重度 | 含义 |
|---|---|---|
| `invalid-run-property` | **error**（退 `1`） | 这个值让 Word **拒绝打开整个包**——「文件可能已经损坏」，COM `Documents.Open` 报 `com_error` `0x800A1529`。这种文件里每一条关系都解析得通、python-docx 也照样能打开，所以本工具其他任何检查都看不见它 |
| `nonstandard-run-property` | **warning**（退 `0`） | 值不合 schema，但 Word **能打开、只是重新解释**。是真事、值得说一句，但不构成拒收 |

实测反例，两个方向各一条：

- `w:highlight="purple"` → **文件打不开**，判 `invalid-run-property`。同档的还有 `DARK_MAGENTA`、`Magenta`、`FF0000`、`darkMagento`、`vertAlign="SUPERSCRIPT"`、`vertAlign="Baseline"`、`effect="blink"`，以及 `<w:highlight/>` / `<w:vertAlign/>` / `<w:effect/>` 这三种**缺 `w:val`** 的写法。
- `w:u="SINGLE"` → **文件能打开**，Word 只是把它重读成 `single`，判 `nonstandard-run-property`。`w:color="#FF0000"`、`w:color="F00"`、`w:color="red"`、`w:color="ZZZZZZ"`、`w:color=""` 同样如此。

**为什么要分两档而不是一档。** 这份报告被 `docx_edit`、`docx_template`、`docx_mathify` 下游消费，`ok: false` 会让它们拒收文档。把 `w:color="red"` 判成 error，等于拒收一份 Word 打得开的文件——**校验器自己成了缺陷**。所以分档才是本轮的设计要点：**严重度由 Word 实测决定，永远不由 schema 决定**，因为 schema 与 Word 在**两个方向上都**会打架：

| 值 | 在 schema 里？ | Word 能打开？ | 判决 |
|---|---|---|---|
| `vertAlign="normal"` | **不在** | **能**（等同 baseline，`Superscript=0 Subscript=0`，存盘时元素被丢掉） | 放行，一条 issue 都不报 |
| `effect="ants"` | **在** | **不能** | `invalid-run-property` |

`normal` 同时也是 `docx_crossref.py` 里「不偏移」的写侧拼法，也就是 `--vert-align` 的默认值：判它 error 等于拒收本仓自己产出的文档。`ants` 是反向陷阱——一个 schema 合法、却被列进合法集合的值，而 Word 根本打不开；这是校验器绝不能产出的那一种结果。两条都用 Word 16.0 COM 实测，一个值一份新造的文档。

`missing-style` 仍然是 **error**，虽然 Word 也打得开，这个区别值得知道：Word 对未知样式 id 不是「重读」，而是在存盘时**换成另一个样式**（实测 `NoSuchStyleXyz` 存回来变成了 `aa`），页面被静默重排。那是内容损失，不是拼法差异。

它是**结构**检查，不判版面。一份文件完全可以返回 `{"ok": true, "issues": []}`，同时有三个半空页和一张越出版心的图——那个判决属于 `--lint`。两个都跑，它们不重叠。

要读 JSON，不要读退出码：`ok: false` + `issues` 是这个工具的正常产出、退 `1`；而路径写错属于输入错误，退 `2`、stderr 一行 `error:`、stdout 为空。这个不对称是故意的，见[退出码](#退出码)。

### 4. 改文本不要用 shell 的字符串替换

PowerShell 的 `-replace`、以及任何 .NET 字符串的 `.Replace()`，**匹配不上时原样返回输入**，什么都不报。脚本在那之后打一行 `saved: …`，只证明调用返回过。同一份文件上它**静默失败至少 6 次**；发现的原因是日志文案不对，不是有任何东西抛了异常。多行文本、中文、带缩进的块最容易匹配不上——一个看不见的空白差异或软换行就够了。

用精确匹配的单点替换：

```bash
python scripts/docx_edit.py replace FILE.docx --find "原文精确片段" --replace "新文本" -o OUT.docx
```

然后读那个计数。`docx_edit.py` 返回 `{"replacements": n, …}`；实测**匹配不上的 `--find` 照样退 `0`、照样写出输出文件、照样声称 `"ok": true`**，只是 `"replacements": 0`。这个计数是「替换真的发生了」的唯一信号，所以**改完必须回读**：对输出跑 `--text` 或 `--runs`，确认旧串没了、新串在。编辑器的精确 edit 或 agent 的 `edit` 工具同样够用；拿正则去刷整个文件不行。

### 检查序列（一块看全）

| 做完这件事 | 跑这个 |
|---|---|
| 存盘（任何写盘脚本） | `docx_preview.py FILE.docx --lint --text`——`sparse_pages` 为空、`issues` 逐条看过 |
| 怀疑内容缺失 | `docx_read.py` 的 `--blocks` / `--text` / `--runs` 三种全跑 |
| 动了结构（表格、图片、页脚、交叉引用） | `docx_validate.py FILE.docx`——`ok: true` |
| 替换或插入过文字 | `docx_read.py --text`，把那段读回来 |
| 准备交付 | 以上全部，再抽一次页面 PNG 亲眼看（`--pages DIR --dpi 150`） |

## 渲染与版面自检

`docx_preview.py` 回答的是工具链其余部分答不出的那个问题：**第 4 页到底对不对？** 它把文档交给无头 LibreOffice 渲染，产出 PDF 与每页一张 PNG，再把「这一页填了多少」量成报告。不用 Word、不走 COM、不需要 Office 授权：

```bash
# 一次转换拿到三样东西
python scripts/docx_preview.py in.docx --pdf out.pdf --pages png/ --lint --json
# 只想看图：只出每页 PNG，dpi 随便调
python scripts/docx_preview.py in.docx --pages png/ --dpi 150
# 只要机器可读的版面报告，不留文件
python scripts/docx_preview.py in.docx --lint --json
```

| 参数 | 作用 |
|---|---|
| `--pdf OUT.pdf` | 把转换出的 PDF 写到这儿（默认写临时文件，用完删掉） |
| `--pages DIR` | 每页一张 `page-%03d.png` 写进 `DIR`，分辨率由 `--dpi` 定（默认 110） |
| `--lint` | 报版面：填充率、半空页、空白页、超宽图、跨页表、孤字末行 |
| `--min-fill RATIO` | 墨迹占比低于这个值就判「稀疏」（默认 `0.40`） |
| `--max-orphan-chars N` | 一页最后一行不超过 N 个字符算孤字（默认 `2`） |
| `--gaps` | 另报绝对空白（mm）：`text_bottom_mm`、`tail_gap_mm`、`head_gap_mm`、`top_margin_mm`，以及 `next_page_lead` 与 `gap_pages` |
| `--max-tail-gap MM` | 正文墨迹停在该页正文区底之上超过这个距离，就算 `gap_page`（默认 `60`） |
| `--skip-last-page-gap` | **不**报末页的页尾空白。默认关闭：末页算不算「半空」是判断题，不是测量结果 |
| `--no-header-band` | 把页眉算成正文墨迹（P18 之前的读法）。只用于复现旧报告 |
| `--timeout SECONDS` | soffice 转换超时（默认 `180`） |
| `--json` / `--text` | 打一个 JSON 对象，或打人类可读的行（默认后者） |

三个能力可任意组合：全传就是一次转换全拿到；至少要给一个。`--lint` 报告带 **8 个 issue code**——常态 7 个，加上 `--gaps` 再多个 `tail-gap`——另有逐页的 `fill_ratio` 数字（它不算 code）。这 8 个是：`sparse_pages`、`blank_pages`、`image_over_width`（图比版心宽）、`table_split`（表跨到了下一页）、`orphan_last`（末行只剩一个字）、`content-width-unknown`（文档没报页宽，图片宽度完全没比对）、`section-margins-differ`（多节边距不一致，几何只取第 1 节）、`tail-gap`。**`table_split` 是启发式、允许误报**：它按表格几何跨页边界比对，所以一张只是恰好在页尾结束的表也可能被点出来。把它读成「这一页值得看一眼」，不是「这里坏了」。`content-width-unknown` 与 `section-margins-differ` 这一对是**降级信号**：它们描述的不是版面，而是「某项测量被跳过或被近似」。一份报了这两条其中之一的文档，它的「干净」比没有这两条时更弱。

### 什么算「正文墨迹」（以及为什么页眉不算）

`--gaps` 回答的是「这页底下有多少白纸」——就是人拿尺子量出来说你那句「空了 8.3 厘米」的那个数。量之前要先去两条带，否则这个数没有意义：

- **页脚**：起始位置低于 `0.88 × 页高` 的块是页码，不是正文；
- **页眉**：结束位置高于**正文区顶**的块是页眉。

页眉这条判据锚在**该文档自己的上边距**上，不是页高的固定比例。这个区别是量出来的、不是口味问题：A4 上 12% 的带结束于 101 pt，而一份上边距 36 pt 的文档，它合法的正文首行结束于 92 pt——按固定比例切，会把每一份紧边距文档的首行悄悄删掉。由 `tests/test_preview_header_band.py::test_a_body_line_near_the_top_is_not_mistaken_for_a_header` 钉住。

这些数字能自己对上，所以你可以**验算而不是相信**：

```
top_margin_mm + head_gap_mm == 正文首行的 y（mm）
```

在 `source-snapshot/test_unpack/test_input.docx` 上：`24.99 + 9.56 = 34.55 mm`，正好是正文首行所在位置（97.9 pt）。切页眉之前，同一页报的是 `head_gap_mm = −9.76`——页眉（`TEST HEADER`，结束于正文区顶之上 12 pt）被当成了正文墨迹。

拿不到上边距时，页眉这条判据是**跳过而不是猜**。这个不对称是故意的：漏掉的页眉会以**负的** `head_gap_mm` 呈现、人会看出来；而被猜出来的切法删掉的正文行，在报告里是看不见的。

依赖与工具链其余部分不同：`docx_preview.py` 要 `PATH` 上有 `soffice`（或在标准 LibreOffice 安装位置），PNG 与 lint 那一半还要 PyMuPDF。缺任何一个都是 exit `1` 加一行 `error:`——属于环境干不了这活，不是参数写错。`fitz` 只在函数内 import、从不放模块顶层，所以光 import 这个模块不需要 PyMuPDF。

> **手工跑 `soffice` 必须加 `-Wait`。** 不用这个脚本、自己驱转换器时，`Start-Process soffice --headless --convert-to pdf …` **不加 `-Wait`** 会立刻返回：退出码是空的，PDF 还没写出来，脚本下一行就去读一个不存在的文件。加 `-Wait`，或者用 `& soffice …` 让 shell 自己等。`docx_preview.py` 没有这个问题，它走 `subprocess.run`，会一直阻塞到子进程退出。

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
| `references/table-geometry.md` | 表格列宽不生效、要决定 `set` 还是 `sync`、`w:zoom` 缺 `w:percent`、用 `keep_together` 把表压在一页、按 `width_frac` 给图定宽 |
| `references/font-scheme.md` | 中英文混排字体规范、`w:rPr` 顺序与去重、`--heading-threshold` 语义、标题为什么默认是黑的以及怎么有意给它上色 |
| `references/merge-and-export.md` | 合并后版心/编号/图片出问题、Word COM 导出细节、PIL 多页 PDF 写法、手工驱 `soffice` 的 `-Wait` 与 `docx_preview.py` |
| `references/revisions-and-comments.md` | 批注与修订的 OOXML 表示与命令行语义 |
| `references/editing-traps.md` | 改既有文档结果悄悄错了：`cell.text =` 把多 run 单元格压平、野 `vertAlign` 让整行变下标、`inline_shapes` 索引漂移、空 `drawing` 让 Word 拒开、WPS/Word 占用报 `PermissionError`、控制台 cp936 乱码 |
| `references/toc.md` | 写静态目录：两遍法（先 `page-probe` 量页码、再写条目）、制表位与点引线、全套坑表、五种页码页脚模板（`zh` / `custom` 占位符） |

## 目录结构

```
scripts/     平铺，按裸模块名 import —— 28 个 .py + 1 个 .ps1 = 29 个文件（27 个 CLI，1 个纯库）
references/  较长的做法笔记，按需加载 —— 7 篇
tools/       Word COM 验收辅助；--page-probe 在这里，即 word_accept.py
tests/       pytest 套件 —— 77 个 .py（76 个 test_*.py + _omml.py 辅助）
```

## 写到哪儿去，以及有什么保护

不同脚本的落点不一样，而这件事直接决定「写错一次」会不会把原件搭进去。

**`-o` 默认覆盖输入。** 所有编辑类 CLI 都收 `-o` / `--output`，不给就等于
「写回你给我的那个文件」：`docx_edit`（九个子命令全收）、`docx_section`、
`docx_crossref`、`docx_toc`、`docx_comments`、`docx_revisions`。help 文本自己写
着 `output path (default: overwrite input)` 或 `(default: overwrite)`。这也正是
「人可能在 Word / WPS 里开着这个文件」时必须显式传 `-o` 的原因——文件被占用会
报 `PermissionError`（见 [`references/editing-traps.md`](references/editing-traps.md)）。

有两个写盘脚本的输出来自**位置参数**、不带默认值——`docx_create`
（`spec.json out.docx`）与 `merge_docs`（`out.docx`）。理由很直接：一个「漏写参数就
覆盖掉别的东西」的写文件命令，永远不是你想要的。两者都直接序列化到那个路径，
都不保留被覆盖内容的副本。

**「原子重写」和「留了备份」不是同一个保证。**

| 脚本 | 机制 | 留 `.bak` 吗 |
|---|---|---|
| `docx_edit`、`docx_section`、`docx_crossref`、`docx_toc`、`docx_comments`、`docx_revisions` | 直接 `doc.save(out)`，python-docx 序列化到目标 | 不留 |
| `docx_template apply --to`、`fix_fonts`、`fix_zoom` | 临时文件 → `shutil.move` / `Path.replace`，中断时原件逐字节不动 | `apply --to` 会写 `<file>.bak`；`fix_fonts` / `fix_zoom` **不写** |
| `docx_mathify --inplace` | `doc.save()` 之后原件被改名成 `<名>.bak.docx` | 会，`.bak.docx` |
| `merge_docs` | 合并结果在内存里拼好，直接写到位置参数 `out` | 不留 |

`fix_zoom` 是原子写那一组里的例外：它**收 `-o`**（默认「就地打补丁」），所以能把
结果写到别处——但换文件依旧走临时文件 + 替换，且照样不留输入副本。`fix_fonts`
**完全没收输出参数**：你给它哪几个文件，它就改哪几个。两者都**崩溃安全**（换文件
是原子的）但**不误操作安全**——参数写错照样得到一个被改写、且没有第二份的文件，
所以文档里那句「原件要紧就自己先复制一份」不是客套。

`apply --to` 与 `--inplace` 的 `.bak` 只有一层——跑两次不会多出 `.bak.1`，而是把
上一次的备份盖掉。

**解包目录是另一种目标。** `table_width.py` 与 `insert_formula.py` 收的是解出来
的**目录**而不是 `.docx`，直接在 XML 分片上改，改完自己打回包。把 `.docx` 递给
它们是用法错误，不会被悄悄当成转换。见[在解包目录上工作](#在解包目录上工作)。

## `{"ok": true}` 到底覆盖了什么，没覆盖什么

`docx_validate` 只回答一个问题：**这个 Word 包是不是结构完好、内部自洽？** `{"ok": true, "issues": []}` 的意思是部件齐全、关系能解析、XML 元素顺序合 schema、没有空的 `<w:drawing/>`、python-docx 打得开。这个结论值得知道，但它**不等于「这份文档看起来是对的」**。

**它是 XML 合法，不是排版健康。** 本工具链自己的产物上量到过两个缺陷，两次都返 `{"ok": true, "issues": []}`：

- 一张**在 16 cm 版心上宽 19.05 cm 的表**——越出页边距，Word 直接画到纸外，而包里没有任何一样东西是畸形的；
- 一个**浮动文本框页脚**，框宽只有 4.56 pt——页码一进位到两位数就被裁，而所有部件依旧合法。

所以 `docx_validate` 绿了是必要条件，不是交付判据。版面相关的问题要另外问：

| 要问的问题 | 用哪条命令 |
|---|---|
| 各列加起来，页面装得下吗？ | `table_width.py check --page-width 16`（越界 exit `1`） |
| 页脚会不会裁字？带不带 PAGE 域？ | `docx_footer.py check` |
| 每页的墨迹落在哪、空了多少？ | `docx_preview.py --lint --gaps` |
| 交叉引用的锚点都指得到吗？ | `docx_crossref.py list` / `check_crossref_anchors` |

一句话记法：**`docx_validate` 读的是包，这几条读的是版面。** 交付前跑 `docx_validate`；想相信它已经做完，还得再跑版面检查。

### `docx_footer check`：那个宽度是**估算**，而且误差是双向的

`page_number_width_pt` 不是量出来的。它不读 `hmtx`、不调排版引擎，是按一套声明的字宽模型算的——全角 1.0 em、ASCII 数字字母 0.5 em、空格 0.25 em、其余半角可打印字符 0.3 em——乘页脚自己的字号。报告里的 `estimate_model` 字段写明用的是哪套模型，所以这个数可以被解释，而不是只能被相信。

**误差往两边走**，这一点要紧，因为读者会把 `overflow: false` 读成「量过了、装得下」：

| 情形 | 模型 | 真实 | 后果 |
|---|---|---|---|
| CJK 卡在框宽上 | 约**低估** 11% | 更宽 | 85 pt 框装 `第 12 页　共 12 页` 实际要 90 pt → 报 `ok`，纸上裁字 |
| 罗马数字 | 约**高估** 80% | 更窄 | 10 pt 框里的 `iii` 只有 7.48 pt，却被报成 error |

所以把这个判定读成**两头可信、边界不可信**：4.56 pt 的框装 88 pt 的行是确凿的发现；而框宽与估算值只差几个点时，应该渲染一页确认（`docx_preview.py --pages`）。`usable_width_pt`（框宽减去内缩）正是为这个比对给的。

还有一件数字不会说的事：文本框是不是**浮动**的。浮动框不参与重排、也不推挤邻居，所以那里裁字就是内容丢失；行内框会换行。报告里的 `shape: "floating-textbox"` 就是把这件事挑明的那一项。

## 更新域：`Fields.Update()` 会连 SEQ 一起重排

文档从 Word 回来、交叉引用是旧的，第一反应是 `doc.Fields.Update()`。那是**全篇、按顺序**的重算，不是「刷新某一个东西」：它按正文顺序走遍每一个域、逐个改写结果——**`SEQ` 计数器也在内**。所以在一份题注本来好好的文档上，一次全量更新反而**会改动**编号（两个题注家族的手工顺序一旦被打散，把表 3 和表 4 弄颠倒的正是这一步）；它还会抹掉那些「手工调过格式、而域指令没把格式钉住」的域结果（`\* MERGEFORMAT`，见上面的上标对照表）。

按意图选做法：

| 想干什么 | 用什么 |
|---|---|
| 只刷新目录 | `doc.TablesOfContents(i).Update()` |
| 只刷新某一个域 | `doc.Fields(i).Update()` |
| 重新分页后刷新页码或目录域 | 更新那个域，而不是整个集合 |
| 只看不重算 | `tools/word_accept.py --no-update-fields` |
| 什么都不做（域本来就带缓存结果） | 直接打开文档 |

`word_accept.py` 的 `--no-update-fields` 就是为最后两行准备的：一个必须观察**文件写了什么**、而不是「Word 会把它变成什么」的探针。

## 退出码

| 码 | 含义 |
|---|---|
| `0` | 成功 |
| `1` | 预期内的失败 —— marker 没命中、未知或重名的书签、文档里没有你要的东西、运行期异常 |
| `2` | 用法错误 —— argparse、spec 里出现未知键或未知 block type、`--widths` 个数多于列数、`--pages` 个数与目录条数不等、互斥 flag 同传（含 `width_frac` 与 `width_mm`）、`--index` / `--insert-at` 越界或负数、输入路径不存在/是个目录/不是 Word 包、`image` 块的 `path` 指向的文件不存在或不是字符串 |
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
- **绝不要用 `cell.text =` 往模板单元格里写。** 它会清空单元格并丢掉全部 run 级属性。用 `docx_edit.py set-cell`（默认 `preserve`）；格子可能不止一个 run 时，先读 `docx_read.py --runs`。
- **`exit 0` 不等于文档是对的。** 输出要读回来——填完模板读 `docx_read.py --runs`，交付前跑 `docx_validate.py`——并且至少用真渲染器打开一次。P13 有四种静默损坏是从全绿的一跑里溜过去的，见 `references/editing-traps.md`。