# docx skill

基于 `python-docx` + `lxml` 的**纯离线** Word 工具集：创建/编辑/读取/校验/批注/修订/模板套用之外，带两项自研能力——**LaTeX → OMML 原生公式链**（不进 Word、不依赖 pandoc/LaTeX 发行版）与**表格几何三处一致**（`tblW` / `gridCol` / `tcW`）。

> 本文是 [README.md](README.md) 的中文镜像：同一批 H2/H3 章节、同一顺序、同一信息量（`### 不止于此` ↔ `### Beyond the basics`，`### 长文档笔记` ↔ `### Reference notes`）。英文版为发布主版，两份如有措辞差异以英文版为准。命令、脚本名、参数名、XML 元素名一律保留原文，不做翻译。

## 为什么还要再造一个 docx 工具链

现成的文档自动化方案通常掉进三个坑之一：

- **要求 Office 常驻**才能跑，命令行、CI、容器里都不成立；
- **把公式画成图片**，看着像公式，实则不可编辑、不可检索、换个字号就糊；
- **把编号与目录写成 Word 域**，文件在有人按下 `F9` 更新域之前是错的——交付物在别人机器上打开就是一片占位文字。

本工具链的选择是：**一堆小而扁平的 CLI 脚本，全部用 `python-docx` + `lxml` 直接改 XML**，核心功能不依赖任何 Office 运行时。唯一需要 Office 的是 PDF 导出（走 Word COM），而那是附加能力，不是主线。

目标是**打开即正确**，而不是"打开再刷一下"：

- **公式是真 OMML 对象**，不是图片，也不是文本近似；整篇批量转或单条插入都行。
- **目录是真段落**，不是域；不装 Word 的读者、文本抽取器、diff 看到的都是完整目录。
- **交叉引用带缓存结果**：`SEQ` 计数 + `REF` 指针各自写好缓存值，`如图3所示` 一打开就是对的，`renumber` 在编号漂移后重新算一遍。
- **编号是真自动编号**：写 `<w:abstractNum>` / `<w:num>` / `<w:numPr>`，而不是套那个不带 `numPr` 的 `List Number` 样式，所以 `[1] … [14]` 真的会渲染出来。
- **表格列宽真的生效**：把 `tblLayout` / `gridCol` / `tcW` 三者的相互作用讲清楚并改对。

### 不止于此

上面那几个硬骨头不是唯一一类"最后一句话往往是『于是我打开 Word 点了一下』"的活儿。另有六项能力覆盖剩下的部分——五项修复，加一条导出通道：

- **`fix_fonts`**——扛得住中英混排的字体统一。`w:rPr` 上有两条硬约束：每个子元素最多出现一次，且 `w:rFonts` 必须是第一个子元素。中文字体要绑到 `eastAsia` 槽，不是 `ascii`，否则 Word 会静默替换掉，不告诉你。
- **`merge_docs`**——把一份文档接在另一份后面，动手检查结果之前值得先知道三条裁定：页面设置取 `base` 的，`src` 的 `sectPr` 被丢掉；`numbering` 按 `w:numId` 取并集；`styles` 按 `w:styleId` 取并集，`base` 已有的样式绝不覆盖。
- **`fix_zoom`**——python-docx 会写出 `<w:zoom w:val="bestFit"/>`，没有 `w:percent`；严格 XSD 校验器直接拒收，Word 的缩放控件打开是一片空白。`docx_create` 已经自己打补丁了，所以这个 CLI 是给别处来的文档用的，一遍就改好。
- **`minidom_compat` + `_bootstrap`**——`defusedxml` 能加固 XML 解析，但多数随包解释器里没有它，于是每个脚本都死在 **import 阶段**。`minidom_compat` 提供 stdlib 兜底，`_bootstrap` 则是把 `sys.path`、这个兜底和 UTF-8 stdio 串起来的那一行。
- **`encoding`**——Windows 下把 stdout/stderr 切到 UTF-8：默认控制台代码页会把中文输出揉成谁也读不了的字节。

还有一项属于**附加能力**而非核心操作。`pdf_export.ps1` 给本工具链产出的文档提供一条不掉公式的 PDF 导出路径：Word 自己的导出器会把 OMML 文本层完整保留，公式仍可搜索、不会退化成图片。**它不是通用的 docx→PDF 转换器**，两个前提必须成立——**文档里真有 OMML 公式**，且**跑在装了 Microsoft Word 的 Windows 上**，因为它驱动的是 Word COM。详见 [Exporting to PDF](SKILL.md#exporting-to-pdf-additional-capability)。

## 不依赖 Word 也能读对的静态目录

`{ TOC \o "1-3" }` 是一个**域**：条目在有人按下 `F9` 之前根本不存在，而手工塞进域里的缓存结果会在按下 `F9` 的瞬间被丢掉。凡是交付物——要转 PDF 的、要进 diff 的、要交给没有 Office 引擎的工具读的——目录都必须是**普通段落**。

`docx_toc.py` 就是这么写的：

```bash
python scripts/docx_toc.py in.docx -o toc1.docx --title 目录
```

每个条目是 `标题<TAB>页码`，用 Word 自己那套 `toc 1/2/3` 样式，右侧制表位停在正文块宽度上、带点引线。制表位宽度取自 `docx_create.content_width_mm(section)`，不许写死：六个实测卷的正文宽度本来就是 170 mm 与 160 mm 两种。

`w:tab` 是这里的关键——页码靠一个真正的制表符推到右边界，而不是靠空格对齐。

`build_static_toc()` 返回的 report 字段：`entries` / `levels` / `styles` / `tab_pos_mm` / `tab_pos_source` / `dropped_entries` / `dropped_titles` / `adopted` / `section_index` / `inserted_at`。重建是幂等且安全的：每个条目段落带一个 `_DocxTocEntry…` 书签，`drop_existing_toc` 只删本模块写过的东西；用户自带的、用内置 `toc 1` 样式写的那份目录不动，除非显式传 `adopt_unmarked=True`。

**页码必须实测，不能猜。** 这是**两遍法**：

1. 第一遍不传 `--pages`——只把条目写进去，页码槽留空（就是上面那条命令的产物，它是要被量的真文件，不是 dry-run）；
2. `python tools/word_accept.py toc1.docx --page-probe 1,3 --json probe.json` 让 Word 排完版，再读回 `paragraph_pages[].displayed_text`。`1,3` 是这份演示文档需要的 1-based Word 段落序号，不是通用常数；每多探一个序号就多两趟 COM 往返；
3. 带着实测页码跑第二遍，**每条目录项正好一个数**——数量对不上算用法错误（exit `2`）：`python scripts/docx_toc.py toc1.docx -o toc-final.docx --pages 1`

页码为什么不能算：`displayed_text` 取自 `Information(WD_ACTIVE_END_ADJUSTED_PAGE)`，是**读者眼睛看到的字形**，不等于物理页序。一份"封面无页码 + 前置小写罗马 + 正文从 1 重开"的文档，物理第 8 页在那一节里印的是 `i`；不同 Word 版本与字体回退会让流式排版落到不同页，任何"按段落序号推页码"的算法都会在真机上错。所以 `source` 字段为 `computed` 的那些行（无 Word 时的估算，`basis: section-start`）**不得当作交付页码**。

两个前提它不会替你造：标题样式必须带 `outline_level`（只是"长得像标题"的样式在 Word 眼里是正文，导航窗格空的、`TOC \o` 也收不到东西）；页码来自对成品文件的第二遍读取。配方、坑单与实测限制见 [references/toc.md](references/toc.md)。

## 公式不是同一种东西

把所有"看起来像公式"的串统统变成公式对象，是最容易做、也最容易做错的做法。本工具链把每个识别出的片段分成**三阶**：

| 片段是什么 | 产出 | 例子 |
|---|---|---|
| **结构式**——含 `\frac`、`\sqrt`、`\int`、`\sum`、`\begin{`、`\left`、`\overline` 等结构命令（`_STRUCTURAL_CMDS`，43 条） | 真 **OMML** 公式对象 | `x^2/16+y^2/9=1`、`v=s/t` |
| **纯上下标**——不含结构命令，但至少有一个 `^` 或 `_` | 普通 run + `w:vertAlign`，即 Word 自己 `Ctrl+Shift+=` 做出来的真上下标，字号可还原 | `a=9.8m/s2`、`10^-7` |
| **免排版**——无结构命令、无 `^`/`_`、恰好一个 `=`、左侧单个 ASCII 字母、右侧数字加可选单位 | 普通文本 run | `s=10km`、`t=2h`、`m=5kg`、`v=3m/s` |

第三阶靠 `docx_mathify.py` 里的 `no_layout_reason()` 判定，单位尾巴是**白名单**而不是"随便什么字母"：每个因子都必须命中 `linear_to_latex.UNITS`。所以 `y=2x`、`n=2t` 仍然是公式（它们是函数关系，不是"2 米""2 吨"），而 `v=s/t` 仍是真分式（`t` 被刻意排除在 `UNITS` 之外）。`x=3` 也落在这一阶，记为 `plain_equation` 而不是 `plain_quantity`。当前的保守边界：`T=25℃`（`UNITS` 里没有 `℃`）与希腊字母量符号 `η=80%` 都还留在公式对象里。

为什么要分阶：公式对象在不需要二维排版的地方**是成本而不是收益**——改一个字符要开公式编辑器、粘到别的程序里掉形、行高被公式引擎撑开。代价是这一阶会丢掉数学斜体（`s=10km` 的 `s` 是正体），这是刻意的取舍：要斜体就给输入写成带上下标的形式（落进第二阶），或在 `docx_create` 的 spec 里显式设 `vert_align`。

两档都可关，开关在 `docx_mathify.py` 的模块级常量：`NO_LAYOUT_NEEDED` 单独管第三阶，`MIXED_SUPERSCRIPT` 是覆盖第二、三阶的总闸（关掉它两档一起停，从而保住"关掉总闸 = 老行为"这条口径）。

公式字体默认 **Times New Roman**，可用 `MATH_FONT` 覆盖（`MATH_FONT="Cambria Math"` 还原 Word 默认观感，`MATH_FONT=""` 完全不注入）。细节见 [references/formula-guide.md](references/formula-guide.md)，三阶的表在 [SKILL.md 的三阶输出表](SKILL.md#three-output-tiers)。

## 环境要求

| 项 | 要求 |
|---|---|
| Python | 3.10+（CI 下限） |
| 依赖 | `python-docx`、`lxml`、`latex2mathml` |
| 测试 | 另需 `pypdfium2`（已列入 `requirements.txt`）与 `pytest`（**未列入**——它在 `requirements.txt` 里是注释行，得自己装：`python -m pip install pytest`） |
| 公式链 | Windows + Microsoft Office：需要 `MML2OMML.XSL`（随 Office 安装，`math_omml.find_xslt()` 自动探测，`MML2OMML_XSL` 可覆盖） |
| PDF 导出 | Windows + Microsoft Word（COM）。**不走 WPS**：WPS 导出带 OMML 的 docx 会丢文本层 |
| 模板 apply（从模板造新文档） | Node.js + `docx` 包 |

Linux / macOS 可用除**公式链**与**PDF 导出**以外的全部功能。

两条边界都在**调用时**才炸，不在安装时预警，所以这里写明：

- **公式链**（`docx_mathify.py` / `insert_formula.py` / 任何走到 `math_omml` 的路径）要 `MML2OMML.XSL`，而 `math_omml.find_xslt()` 只探测 Windows 的 Office 安装位置（`%ProgramFiles%`、`%ProgramFiles(x86)%`、`%LOCALAPPDATA%`、写死的 `C:\Program Files…`），Linux / macOS 上必然抛 `FileNotFoundError`——除非自己放一份并设 `MML2OMML_XSL` 指过去。这条链目前只在 Windows 上实测过。
- **PDF 导出**（`pdf_export.ps1`）走 Word COM，天然绑 Windows + 本机 Word。

其余全部（创建 / 读取 / 编辑 / 校验 / 批注 / 修订 / 编号 / 交叉引用 / 静态目录 / 分节 / 表格几何 / 字体 / 合并 / 模板 / `docx_pack`）是纯 `python-docx` + `lxml`，**机器上完全不用装 Office**。

旧式二进制 `.doc` 不支持——先用 Word 另存为 `.docx`，或在外部用 LibreOffice 转换。

## 安装

面向 agent（Claude Code、Cursor、Codex、Copilot、Gemini CLI 等）：

```bash
npx skills add AureliaRen/docx-skill
```

面向人工部署：

```bash
git clone https://github.com/AureliaRen/docx-skill.git
cd docx-skill
python -m pip install -r requirements.txt
```

Python **3.10** 或更新。`requirements.txt` 里只有 `python-docx`、`lxml`、`latex2mathml` 三条必需项，外加只有 PDF 文本层断言才用得到的 `pypdfium2`。`pytest` **不在其中**：第 19 行是 `# pytest>=7.4`，一条注释，因为它是纯开发工具。所以 `pip install -r requirements.txt` 之后跑 `python -m pytest` 会得到 `No module named pytest`，顺手一起装——`ci.yml` 就是这么装的：

```powershell
python -m pip install -r requirements.txt pytest
```

两项可选依赖：

| 能力 | 额外要求 |
|---|---|
| 公式转换 | `latex2mathml`（纯 Python，无编译依赖）+ 本机 Microsoft Office 提供的 `MML2OMML.XSL` |
| 模板 apply | Node.js + `docx` 包 |

公式转换复用本机 Office 自带的那份样式表，不重新发明一遍 MathML→OMML 的映射。装在非常规位置时用环境变量指过去：

```powershell
$env:MML2OMML_XSL = 'D:\Office\root\Office16\MML2OMML.XSL'
```

`docx_mathify.py` 与 `insert_formula.py` 共用同一个探测函数，只有公式功能才需要它——其余脚本在没有 Office 的机器上照跑。

## 典型用法

18 个带 argparse 表面的 CLI，先看 `--help` 再动手；另有 7 个模块没有 argparse 表面（只能 `import`，或裸跑看一行自述 / 跑自检），`_bootstrap.py` 是唯一的纯库——逐文件清单见下一节「脚本清单」的末列：

```bash
python scripts/docx_create.py --help
python scripts/docx_read.py --help
```

`docx_create.py` 吃两个**位置参数**：`python scripts/docx_create.py <spec> <output>`。它没有 `--spec` / `--output` 开关，传了是用法错误，不会被静默忽略：

```text
docx_create.py: error: unrecognized arguments: --spec --output   (exit 2)
```

`table_width.py` 与 `insert_formula.py` 操作的是**解包目录**（把一个 `.docx` 展成目录直接改 XML），先解包、改完再打包：

```python
from docx_pack import unpack, pack
unpack("in.docx", "unpacked/")
# ... 改 unpacked/ 下的 XML ...
pack("unpacked/", "out.docx", original="in.docx")
```

上面这段是 **Python**，不是 shell：`docx_pack` 在 `scripts/` 里，所以在仓根直接贴进解释器会 `ModuleNotFoundError`。要么先 `cd scripts`，要么先把目录塞进 `sys.path`——一条命令版是 `python -c "import sys; sys.path.insert(0,'scripts'); from docx_pack import unpack; unpack('in.docx','unpacked/')"`。

`pack` 沿用原件的条目顺序与压缩方式，所以"什么都没改"的往返产出等价归档。

`unpack` 把归档当**不可信输入**：条目名里的 `../`、绝对路径、反斜杠、以及 Windows 盘符形式（如 `D:/x.txt`）一律拒绝，不让压缩包写到解包目录之外，命中即抛 `UnsafeArchiveError`。

五个常用动作：

```bash
# 1. 整篇线性公式批量转原生 OMML，并出转换报告
python scripts/docx_mathify.py in.docx -o out.docx --report conv.md --tex formulas.tex

# 2. 插单条行间公式（锚定在 "marker" 段落之后）
python scripts/insert_formula.py unpacked/ "\frac{-b+\sqrt{b^2-4ac}}{2a}" --display --after "marker"

# 3. 表格几何：读现状 → 按 cm 写死四列
python scripts/table_width.py report unpacked/
python scripts/table_width.py set unpacked/ --table 1 --widths 2.4,4.2,2.4,5.6

# 4. 合并并导 PDF
python scripts/merge_docs.py base.docx src.docx merged.docx
pwsh scripts/pdf_export.ps1 -DocxPath merged.docx -PdfPath merged.pdf

# 5. 占位符填充（--strict：有 token 没填就报错，不静默放过）
python scripts/docx_template.py fill template.docx values.json out.docx --strict
```

第 2、3 步吃的都是**解包目录**，先按上面那条 `python -c "…unpack('in.docx','unpacked/')"` 造一个；`--after "marker"` 要求文档里真有一个含 `marker` 的段落，第 3 步的四列宽度也要求那张表正好四列。

最小往返（造 → 读 → 转公式）。spec 直接用仓里的，不用在命令行里拼 JSON：

```bash
# 造——examples/formulas.json 是仓里自带的示例，含 5 条可转公式
python scripts/docx_create.py examples/formulas.json report.docx

# 读
python scripts/docx_read.py report.docx --structure

# 转公式
python scripts/docx_mathify.py report.docx -o report-math.docx --report conv.md
```

`examples/basic.json` 是覆盖面更广的那份（十种常用块 + 表格 + 分页），想看块类型先看它；`examples/formulas.json` 才是该喂给 `docx_mathify` 的那份。

**在 PowerShell 里写自己的 spec**：bash 的 heredoc `cat > spec.json <<'JSON' … JSON` 在 `powershell.exe` 与 `pwsh` 下都直接解析失败——`ParserError: 重定向运算符后缺少文件规范`，exit `1`。改用 here-string：

```powershell
$spec = @'
{"blocks": [
  {"type": "heading",   "text": "Report", "level": 1},
  {"type": "paragraph", "text": "Ohm's law: U=IR"}
]}
'@
Set-Content -Path spec.json -Value $spec -Encoding utf8
python scripts/docx_create.py spec.json report.docx
```

完整命令表在 [SKILL.md](SKILL.md)（中文版 [SKILL.zh-CN.md](SKILL.zh-CN.md)）。

还有一个面向 agent 的往返：结构性结果一律走 **stdout 的 JSON**——`docx_create` / `docx_read` / `docx_edit` / `docx_validate` / `docx_comments` / `docx_revisions` / `docx_section` / `docx_numbering` / `docx_crossref` / `docx_toc`；其余出人读文本（`docx_mathify` / `table_width` / `insert_formula` / `docx_template`，`fill` 除外，它出 JSON）。诊断信息一律走 stderr。退出码别只判「是不是 0」，按四档分：

| 码 | 含义 |
|---|---|
| `0` | 成功 |
| `1` | 预期内失败——锚点段没找到、spec 不合法 |
| `2` | 用法错误——argparse 拦下、某个必需的模式开关没给、`--index` / `--insert-at` 越界，或输入路径不存在、是目录、不是 Word 包 |
| `3` | 部分成功——`docx_mathify` 把文档写出去了，但至少还有一条公式是线性文本 |

`3` 出自 `docx_mathify`，而 `docx_template apply --fix-formulas` 会把它原样透传——因为那条通道就是就地跑一遍 mathify；`--no-fix-formulas` 压根不调 mathify。两种情况都是刻意的：产物有效、绝大多数已转换，所以交付前要看一眼退出码。

用法错误只往 stderr 写一行、stdout 留空、什么都不碰——输入文件逐字节不变。这条在 `docx_read` 上最要紧：它每一种成功模式都出 JSON，所以缺文件、给的是目录、不是 ZIP、是 ZIP 但不是 Word 包，一律 exit `2` 且 stdout **一个字节都没有**，而不是回一份会被每个 `json.loads(stdout)` 调用方读成成功的失败 JSON。

## 脚本清单（`scripts/`，27 个文件）

`scripts/` 下 **26 个 `.py` + 1 个 `.ps1`，共 27 个文件**，全部平铺、无子包、无 `__init__.py`。分类只用一条判据，且 `ci.yml` 是机械推导的、不靠这段文字自觉——两条规则取并集（`_*` 前缀 **或** 没有 `if __name__ == "__main__"` 守卫）即判为**纯库**：

| 桶 | 数量 | 判据 |
|---|---|---|
| **CLI** | **25** | 带 `__main__` 守卫，能直接从命令行跑 |
| **纯库** | **1** | `_bootstrap.py` —— 私有模块，在解析参数之前就被 import，永远不单独运行 |

这 25 个 CLI 里，**18 个有 argparse 表面**（`--help` 打印 `usage: …`）。另外 7 个有 `__main__` 但没有 argparse：6 个只打印一行「是库模块，import 它」自述（`docx_pack` / `docx_utils` / `encoding` / `linear_to_latex` / `math_omml` / `minidom_compat`），`fix_phasor_dot` 跑的是自检。它们与 `_bootstrap` 一样按裸模块名 import。每个命令的完整参数见 [SKILL.zh-CN.md](SKILL.zh-CN.md)。

**CLI 工具（18 个带 argparse 的 `.py`）**

| 文件 | 作用 |
|---|---|
| `docx_create.py` | 按 JSON spec 生成 `.docx`（`page` / `styles` / `header` / `footer` / `blocks`） |
| `docx_read.py` | 读回文档：`--text` / `--structure` / `--styles` / `--images DIR` / `--revisions` / `--blocks` / `--md` |
| `docx_edit.py` | `replace` / `set-cell` / `insert` / `delete` / `style` / `normalize` / `toc` / `page-numbers` |
| `docx_validate.py` | 包级健康检查，`ok:false` 时 exit 1 |
| `docx_comments.py` | 批注 `list` / `add --target` / `delete --id`（**id 从 `0` 起**） |
| `docx_revisions.py` | 修订 `list` / `accept-all` / `reject-all` / `accept` / `reject` |
| `docx_common.py` | 共享段落遍历：`iter_all_paragraphs` / `iter_part_roots` / `replace_in_paragraph`；`--api` 打印导出面 |
| `docx_template.py` | `extract` / `apply`（`-o` 需 node + docx-js，`--to` 打回现有文档，`--fix-formulas` 透传 mathify 的退出码）/ `fill`（`{{token}}`） |
| `docx_numbering.py` | 自动编号 `list` / `define` / `apply` / `detach` / `restart` |
| `docx_crossref.py` | 交叉引用与题注 `caption` / `ref` / `bookmark` / `renumber` / `list` |
| `docx_toc.py` | 静态目录 `build_static_toc` / `collect_outline` / `drop_existing_toc` |
| `docx_section.py` | 分节 `list` / `break` / `pagenum` / `pagenums` / `header` / `footer` / `cols` |
| `docx_mathify.py` | 批量把线性公式换成 OMML（`--report` / `--tex` / `--dry-run` / `--no-merged-script-fix`）；`--dry-run` 与 `-o` 同传是 exit `2` |
| `insert_formula.py` | 解包目录插单条 LaTeX 公式（`--after` / `--before` / `--display` / `--dump`） |
| `table_width.py` | 表宽 `report` / `set --widths` / `sync` / `auto --content-width` |
| `fix_fonts.py` | 字体统一：样式层 + run 层 + 手动放大标题（`--heading-threshold`） |
| `fix_zoom.py` | 给 `<w:zoom>` 补 schema 强制的 `w:percent`（`--percent`） |
| `merge_docs.py` | 合并两份 docx（样式/编号/图片关系/sectPr，`base src out`） |

第 27 个文件不是 `.py`：`pdf_export.ps1` 是 PowerShell 导出通道——Word COM、`ExportAsFixedFormat(17)`、以 `ReadOnly` 打开（`-DocxPath` / `-PdfPath`）。详见 [导出 PDF](SKILL.zh-CN.md#pdf-导出附加能力)。

**库与自述模块（7 个）**——按裸模块名 import；前六个带一行 `__main__` 自述，`fix_phasor_dot` 跑的是自检。`_bootstrap.py` 是唯一的纯库，单独列在最后。

| 模块 | 提供什么 |
|---|---|
| `docx_pack.py` | `unpack` / `pack` —— 把 `.docx` 解成目录再打回去，还原原始条目顺序与逐条目压缩方式。`table_width.py` 与 `insert_formula.py` 吃的就是它产出的目录，虽然两者都没 import 它。 |
| `docx_utils.py` | `wtag` / `insert_ordered` / `fix_child_order` / `cm_to_dxa` / `FONT_TNR`、`FONT_HEITI`、`FONT_SONG` |
| `linear_to_latex.py` | `find_math_spans` / `linear_to_latex` / `repair_merged_scripts`，外加 `UNITS` 白名单 |
| `math_omml.py` | LaTeX → MathML → OMML：`latex_to_omml_lxml` / `assert_legal` / `apply_math_font` |
| `fix_phasor_dot.py` | `m:limUpp` → `m:acc`，修相量点浮大；`python scripts/fix_phasor_dot.py` 跑自检 |
| `encoding.py` | `_fix_stdio()` —— Windows 下把 stdout 与 stderr 切到 UTF-8 |
| `minidom_compat.py` | `defusedxml.minidom`，缺包时退回 stdlib |

还有唯一一个**纯库**文件：

| 模块 | 提供什么 |
|---|---|
| `_bootstrap.py` | 平铺布局自举：`sys.path` 自定位加 minidom / UTF-8 兜底。没有 `__main__`、没有 argparse——其余每个脚本都在解析参数之前 import 它。 |

## 仓内目录与 import 纪律（Layout）

```
scripts/    所有 CLI，平铺，按裸模块名 import —— 27 个文件（25 个 CLI + 1 个纯库 + 1 个 .ps1）
tools/      word_accept.py 等验收工具；**`--page-probe` 只在这里**
references/ 长文档笔记，按需阅读，不进 import 路径
tests/      pytest 套件
examples/   示例 spec（basic.json / formulas.json）
```

`tools/` 只有一个大件 `word_accept.py`，靠 Word COM 接受待处理修订、并报出每个段落实际落在第几页。**`--page-probe` 是它的开关，`docx_toc.py` 自己没有探测参数**——先用 `word_accept.py --page-probe` 量，再把量到的页码通过 `docx_toc.py --pages` 传回去。

所有脚本共用一条 import 纪律：**先把自身所在目录塞进 `sys.path`，再 import 同级模块**。于是 `python scripts/x.py`、在 `scripts/` 里 `python x.py`、以及从别的目录用绝对路径调用，三种方式都能跑——没有一条 `__init__.py` 链要维护，也就没有一处会写错。

### 长文档笔记（Reference notes）

六篇长文档笔记按需查（都在 `references/`）：

| 笔记 | 什么时候读 |
|---|---|
| [references/toc.md](references/toc.md) | 静态目录的两遍法、页码实测、书签名三条硬要求 |
| [references/formula-guide.md](references/formula-guide.md) | 三阶判定细节、源文本踩坑、`MATH_FONT` 与两条已修缺陷 |
| [references/table-geometry.md](references/table-geometry.md) | `tblLayout` / `gridCol` / `tcW` 三处一致与实测对照 |
| [references/revisions-and-comments.md](references/revisions-and-comments.md) | 批注锚点与修订接受/拒绝的边界 |
| [references/font-scheme.md](references/font-scheme.md) | 样式层 + run 层字体统一、标题放大阈值的判定 |
| [references/merge-and-export.md](references/merge-and-export.md) | 两份 docx 合并与 PDF 导出的注意事项 |

## 测试

```bash
python -X utf8 -m pytest tests -q
```

`tests/` 用 `subprocess` 调 argparse CLI，fixture 现场生成、产物写临时目录。依赖 `latex2mathml` **加** `MML2OMML.XSL` 的用例（`test_formula` / `test_units` / `test_mixed_superscript`）**自己 skip**，并写明原因 `needs latex2mathml + Microsoft MML2OMML.XSL (local Office install)`——所以没有 Office 的机器得到的是**绿跑 + 更大的 skip 数**，不是红灯。

同一个 commit、两台机器，数字不同只因为 skip 不同：

| 环境 | 结果 |
|---|---|
| 开发机，本机装了 Office | **1576 passed / 239 skipped / 0 failed**（collected 1815） |
| 模拟 runner，整机没有 Office（`MML2OMML_XSL` 指向一个不存在的文件，**且必须在仓库的 clone 里跑**——开发树的上层目录有 `_smoke/`，会让其中一条用例从 skip 变成 pass）——`ci.yml` 卡的就是这一行 | **1440 passed / 375 skipped / 0 failed**（collected 1815） |

差的就是那 135 条：需要 `latex2mathml` **加** `MML2OMML.XSL` 的用例在第一行真跑、在第二行 skip。两组数字都不是拿来改常数追平的——`ci.yml` 的门禁是 `passed >= 1440` 与 `skipped <= collected / 4`。

`tests/` 里是 46 个 `test_*.py` 模块加一个共享助手 `_omml.py`。`pypdfium2` 用于导出 PDF 后断言文本层，`pytest` 是测试运行器，两者都不是运行期依赖。没有 Word 的机器上，涉及 COM 的用例同样自动 skip。

## 授权

MIT，见 [LICENSE](LICENSE)。第三方来源与运行期组件单独记在 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)。

本项目的骨架来自 [NousResearch/hermes-agent](https://github.com/NousResearch/hermes-agent)（MIT，`skills/productivity/microsoft-office`）里的 clean-room 实现：`docx_create` / `docx_edit` / `docx_read` / `docx_validate` / `docx_comments` / `docx_revisions` / `docx_template` / `docx_common` / `docx_numbering` / `docx_crossref` / `docx_toc` / `docx_section` 这 12 个文件由它衍生（**判据是各文件头部的 `# MIT License. Part of the Hermes docx skill.` 标记，那是权威来源标记**），本地改动限于 `sys.path` 自举头、UTF-8 stdio 与若干可移植性/缺陷修复（`docx_common.py` 另加 `--api` 自述入口与替换迭代预算，防替换值本身含查找串时死循环）。上游 MIT 授权对它们原样适用。

**自研增量**（头部不带上面那个 Hermes 标记的那批）：`linear_to_latex.py`、`math_omml.py`、`fix_phasor_dot.py`、`docx_mathify.py`、`insert_formula.py`（公式链）、`table_width.py`、`fix_zoom.py`、`fix_fonts.py`、`merge_docs.py`、`pdf_export.ps1`（本地移植的 6 项能力即 `fix_fonts` / `merge_docs` / `pdf_export` / `fix_zoom` / `minidom_compat` / `encoding` 所在的那批），以及 `docx_pack.py`。

**不随本仓库分发**的组件：`MML2OMML.XSL` 是 Microsoft Office 的一部分，运行时从本机 Office 读取；`pdf_export.ps1` 通过 COM 驱动本机 Word，脚本本身不含微软代码。`latex2mathml`、`python-docx`、`lxml` 也由用户自行安装，不 vendor 进来。上游早期版本的 `LICENSE.txt` 属专有条款，未随本仓库分发。

## 致谢

感谢 NousResearch 的 [hermes-agent](https://github.com/NousResearch/hermes-agent)（MIT）：基础 CRUD、批注、修订与模板 CLI 由那份 clean-room 实现改写而来，它本身是照功能规格对着 `python-docx` 一行行写出来的。逐文件的来源归属与本地改动清单见 [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)。
