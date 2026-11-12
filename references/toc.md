# 静态目录（TOC）：把目录写成真正的段落

> 本文覆盖 `scripts/docx_toc.py`（`build_static_toc` 一族）与它依赖的两块前置能力：
> `scripts/docx_create.py` 的 `outline_level` / `ensure_toc_styles` / `content_width_mm`，
> 以及 `tools/word_accept.py` 的 `--page-probe`（唯一可信的页码来源）。
> 所有数字都是本机 Word 16.0 实测，不是推断。

## 1. 两条路线，先选一条

| | `{"type":"toc"}`（TOC **域**） | `toc.build_static_toc()`（**静态目录**） |
|---|---|---|
| 文件里存的是什么 | 一个 `{ TOC \o "1-3" \h \z \u }` 域 + 一句占位结果 | 真正的段落：`第一章　静电场<w:tab/>1` |
| 谁生成内容 | Word / LibreOffice 按 F9（更新域）现算 | 本模块，生成时就算好 |
| 不装 Office 的读者看到什么 | 占位句（实测：`Table of contents - open in Word/LibreOffice and update fields to populate.`） | 完整目录 |
| `docx_read.py --text` / 文本抽取 / diff | 只有占位句 | 完整目录 |
| 正文改动之后 | F9 一下自动跟上 | **页码写死在文件里，必须重新生成** |
| 条目可点击 | 是（域生成的 `HYPERLINK \l "_Toc…"`） | 否（纯文本行，没有书签锚点） |
| 样式由谁决定 | Word 用它自己那套 `toc 1/2/3` | 本模块写 `toc 1/2/3`（见 §2.2） |

选型口径：

- **交付物必须在任何环境里"打开即正确"** → 静态目录。
- **文档还会被人在 Word 里反复改，且读者一定用 Word** → 域更省事。
- **两者都要** → 只能选一个位置放。同一份文档里同时存在域和静态目录，按 F9 之后读者会看到
  两份目录。实测（`both.py`，同一份文档里既有静态 `toc 1` 行、又有一个 TOC 域）：

```
before F9: ['目录', '第一章　静电场\t1', '1.1　库仑定律\t1',
            '第一章　静电场', '1.1　库仑定律', 'placeholder']
after  F9: ['目录', '第一章　静电场\t1', '1.1　库仑定律\t1',
            '第一章　静电场', '1.1　库仑定律',
            '第一章　静电场\t1', '1.1　库仑定律\t1', '']      ← 第二份目录
```

域那一路也不会把静态目录"吃掉"或"合并掉"：它按大纲级别收集，我们写的 `toc 1` 行不是标题，
所以两边各自列一份。挑一条路走到头。

### 真风险：域的预填缓存会被 F9 丢掉

`docx_create` 的每个域都带缓存结果（`begin / instrText / separate / result / end`），
这样文件不重算也能读对。**TOC 域是例外，缓存保不住。**

实测（`f9_risk.py`，一个手写缓存的 TOC 域，缓存内容是
`第一章 静电场……9 ｜ 1.1 库仑定律……9`）：

```
A. Word 没碰过它
   body: ['第一章　静电场', '1.1　库仑定律', '第一章 静电场……9 ｜ 1.1 库仑定律……9']
   field cache: [('TOC \\o "1-3" \\h \\z \\u', '第一章 静电场……9 ｜ 1.1 库仑定律……9')]

B. Word 打开 → Fields.Update()（= F9）→ 保存
   body: ['第一章　静电场', '1.1　库仑定律', '第一章　静电场\t1', '1.1　库仑定律\t1', '']
   field cache: [('PAGEREF _Toc244978928 \\h', '1'), ('PAGEREF _Toc244978929 \\h', '1'),
                 ('TOC \\o "1-3" \\h \\z \\u', '第一章 静电场1.1 库仑定律')]

our fake text survived: False
```

也就是说：**给 TOC 域预填一份手工目录是无效投入**——它只在"没人按 F9"的时候存在，
按一次就没了，而且换来的假页码会误导任何按文本处理文件的下游。
要么接受域（读者自己按 F9），要么走静态目录。别混。

## 2. 前置能力

### 2.1 `outline_level`：没有它，目录必然是空的

TOC 域按 `\o "1-3"` 收**大纲级别**，本模块的 `collect_outline()` 判断"这是不是标题"也
只看 `docx_read.heading_level()`。这两条路的前提都是样式上带 `<w:outlineLvl>`。

模板自带的 `卷标题` / `节标题` 如果只是"看起来像标题"（黑体、加粗、居中）而不带
`outlineLvl`，Word 眼里它就是正文。实测三种写法（同一份文档，只差 `outline_level`，
`{"type":"toc"}` 域 + F9）：

| `styles[].outline_level` | 写进 XML 的值 | TOC 域 F9 之后 | `collect_outline()` |
|---|---|---|---|
| `1` | `w:val="0"` | `第一章 静电场 1 1.1 库仑定律 1` | 收得到 |
| 省略 | 什么都不写 | `未找到目录项。` | `[]` |
| `0` | `w:val="9"`（Word 的正文标记） | `未找到目录项。` | `[]` |

所以：

- **`1`..`9` 是 1-based 的大纲级别**，`1` 写成 `w:outlineLvl val="0"`（Word 内部 0-based）。
- **`0` 是"显式正文"**，写成 `val="9"`，用来明确说"这个样式永远不要被收集"。
  它和"省略这个键"在行为上等价，但意图不同——省略是没定义，`0` 是定义成正文。
- 键写在 `styles[]` 条目上，作用于**样式**，不是段落。段落继承即可，不要在段落上再写一遍。

```json
{"styles": [
  {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
   "bold": true, "outline_level": 1},
  {"name": "节标题", "base": "Normal", "font": "黑体", "size_pt": 14,
   "bold": true, "outline_level": 2}
]}
```

### 2.2 `toc 1/2/3` 样式：Identity 与外观是两件事

两个函数，分工不同，**两个都要叫**：

| 函数 | 归谁 | 做什么 | 实测效果 |
|---|---|---|---|
| `docx_toc.ensure_toc_style(doc, level)` | `docx_toc` | 只保证样式**存在**、`styleId` 钉成 `TOC1`、去掉 `w:customStyle`；样式体是空的（based on `Normal`，无字体无字号） | Word 认得这是它的内置 `TOC 1`，但行距用 Word 自己的潜默认（13.8 pt） |
| `docx_create.ensure_toc_styles(doc, levels, indent_chars)` | `docx_create` | 把整个定义补齐：`basedOn=Normal` / `next=Normal` / `uiPriority=39` / `unhideWhenUsed` / `qFormat` / 1.5 倍行距（`w:line=360`）/ 逐级缩进（`0 / 2 / 4` 字符，同时写 `w:left` 与 `w:leftChars`） | 生成的 `TOC 1` 行 18.0 pt 行距、`TOC 2` 左缩进 22.0 pt，即我们的定义生效 |

`build_static_toc` 内部会对每个用到的级别调一次 `ensure_toc_style`（只为拿 `styleId`）。
**要外观就必须先自己调 `ensure_toc_styles`**：

```python
import docx_create, docx_toc
docx_create.ensure_toc_styles(doc, levels=(1, 2), indent_chars=(0, 2))
docx_toc.build_static_toc(doc, levels=(1, 2), title="目录", insert_at=1)
```

`ensure_toc_styles` 是幂等的：已有同名样式就地补全（`created` 为空、`updated` 列出名字），
不会造出第二个同名 `w:style`——那会让 Word 报文件损坏。

> 关于 `styleId`：`docx_toc` 钉 `TOC1`（大写），`docx_create` 钉 `toc1`（小写）。两种写法
> Word 都认（实测：`TOC1` 定义下生成的 `TOC 1` 行吃到了自定义行距 18.0 pt，`toc1` 同样
> 18.0 pt）。所以这不是缺陷，只是两个模块的写法没统一。

## 3. 页码从哪来：`--page-probe`

静态目录的页码只有一处可信来源：**让 Word 排版，然后问它"这一页上印的是几"**。

```powershell
python tools/word_accept.py paper.docx --page-probe 7,10,13,16 --quiet
```

在仓库根执行。I,J,... 换成你文档里真实的标题段落序号（§8.2 就是怎么算出来的）。序号越界
不会报错，只会降级：在只有 3 段的文档上跑这一行，四个序号全越界，
`paragraph_pages_source` 变成 `"none"`、`displayed_text` 全 `null`，`warnings` 里逐条
`page_probe_indices: paragraph 7 out of range 1..3`；部分越界才是 `"word-partial"`。

`--page-probe I,J,...` 收的是 **Word 的段落序号（1-based）**，不是 XML 的 `w:p` 序号——
Word 给每一行表格额外算一个段落，两者天然不等。每个序号两次 COM 往返，所以是 opt-in，
上限 `PAGE_PROBE_BUDGET = 60` 条。

返回 `paragraph_pages[]`，每行两个不同的页码：

| 键 | 取自 | 含义 |
|---|---|---|
| `physical` | `Range.Information(WD_ACTIVE_END_PAGE)`（= 3） | **物理页序**：从文档第一张纸开始数，忽略所有分节重启与格式 |
| `displayed_number` | `Range.Information(WD_ACTIVE_END_ADJUSTED_PAGE)`（= 1） | **版面页码**：计数器被 `w:pgNumType` 调整之后的那个整数 |
| `displayed_text` | `page_glyph(fmt, displayed_number)` | 读者眼睛看到的**字形**：`i` / `３` / `③` / `- 3 -` / `""` |
| `section` | Word 自己的 `Sections(i).Range` 区间比对 | 这个段落在第几节（1-based） |
| `page_format` | `section_specs[section-1].fmt` | 该节的计数器族 |

### 物理页序 ≠ 版面页码：写错就是错

一份"封面 + 前置(小写罗马) + 正文(重新从 1 开始)"的文档，实测四行：

```
idx  sec  physical  displayed_number  displayed_text  fmt
9    2    3         1                 'i'             lowerRoman
12   2    4         2                 'ii'            lowerRoman
14   3    5         1                 '1'             decimal
17   3    6         2                 '2'             decimal
```

物理第 3 页在印 `i`，物理第 5 页在印 `1`。目录要写 `['i', 'ii', '1', '2']`；
写成 `[3, 4, 5, 6]` **就是错的**，读者在纸上找不到那个页码。

同一个坑在六卷实测里出现过两次：math 卷 5 节、最后一节 `lowerRoman start=1`，
物理第 8 页在那一节里印的是 `i`；physics 卷物理第 6 页印的是 `v`。
按物理页序回填的目录，这两处都是错的。

**`displayed_text` 才是要回填的东西。**

### 一个必须知道的接线细节

`displayed_text` 的字形由该节的 `fmt` 决定。直接调 `probe_word(...)` 而不传
`section_formats=`，`fmt` 就是 `None`，`page_glyph(None, n)` 退化成 `str(n)`——
`i` 会变成 `1`，静默且难查。实测同一份文件：

```
不传 section_formats:  9  ->  physical=3  displayed_number=1  '1'  page_format=None
传 section_formats:    9  ->  physical=3  displayed_number=1  'i'  page_format='lowerRoman'
```

**走 CLI**（`tools/word_accept.py`）就不必操心：`build_report()` 会自己把
`section_specs[].fmt` 列表传进去。要在 Python 里直接调，就自己传：

```python
import sys; sys.path.insert(0, "tools")
import word_accept as wa
xml = wa.probe_xml(path)
rep = wa.probe_word(path, page_probe_indices=indices,
                    section_formats=[s["fmt"] for s in xml["section_specs"]])
pages = [r["displayed_text"] for r in rep["paragraph_pages"]]
```

## 4. `source` 纪律：`computed` 不是页码

`paragraph_pages_source` 三个值：

| 值 | 含义 | 能不能当交付页码 |
|---|---|---|
| `word` | 每个请求的段落都拿到了 Word 实测的 `displayed_number` | **能** |
| `word-partial` | 一部分拿到了，一部分没拿到（越界、COM 失败） | 只能逐行看 `displayed_number is not None` 的那几行，其余不能用 |
| `none` | 一条都没测到（没要 `--page-probe`、或 Word 不可用） | 不能 |

没有 Word 时 `build_report()` 会退到 `estimated_paragraph_pages()`，`source` 变成
`computed`。这些行是从 `w:pgNumType` 加节的段落区间算出来的：

```
{'index': 9, 'section': 2, 'page_format': 'lowerRoman', 'physical': None,
 'displayed_number': 1, 'displayed_text': 'i', 'section_last_page': None,
 'exact': False, 'basis': 'section-start', 'index_space': 'xml', 'source': 'computed'}
```

`basis: "section-start"` 说的就是它的真实含义：**这是该节第一页的页码，不是这个段落的页码**。
它不知道段内位置（那要排版），而且继承 XML 的索引空间（有表格就与 Word 差一段/行）。
`warnings` 里会同时挂一句 `computed 不得当作交付页码`。

**用途**：说明"某段落在哪一节、那一节用哪种数字"，或者做结构断言。
**不能用作**：`build_static_toc(pages=[...])` 的输入。

## 5. 书签名：三条硬要求，全是静默失败

`docx_toc` 给每个条目段落打一个 `_DocxTocEntry0001` 形状的书签（`ENTRY_MARK_PREFIX`），
用来识别"这是我写的"。名字由 `docx_crossref.next_toc_name()` 生成，天然满足下面三条。
自己写书签（`bookmark` 键、`caption` 的 `bookmark`、交叉引用目标）时必须自己守住。

| 要求 | 违反后 Word 的行为 | 违反后 LibreOffice 的行为 | 检查函数 |
|---|---|---|---|
| **唯一** | **静默合并**：两个 `bookmarkStart` 都还在 XML 里，Word 只留一个（留后者），指向后一个的 `REF`/`PAGEREF` 跳到前一个的位置 | 把第二个**改名**成 `<名字>_副本_1`，所有用原名写的域全死 | `validate_bookmark_uniqueness(doc)`、`check_crossref_anchors()["duplicate_names"]` |
| **无空格** | 保存时把空格换成 `_`，`REF _Ref图 1 \h` 被读成"REF _Ref图"加一个字面量 1，域废 | 同样换成 `_` | `validate_bookmark_name()` |
| **≤ 40 字符** | 静默截断到 40 | 静默截断到 40 | `validate_bookmark_name()` |

三条都是**静默**的：XML 看着没问题，两个应用都不报错，坏在读者手里。
所以 `docx_create.main()` 把 `invalid_names` / `duplicate_names` 当**硬失败**
（`ok: false`、退出码 1、不写文件）。`check_crossref_anchors()` 的返回值里这两个键与非空
`missing` 同级对待。

### 字符集：只拦控制字符，别自己加限制

很容易想当然写成"只许字母数字下划线"，**那是错的**。Word 16.0 实测 22 个名字，**21 个原样
存活并可解析**：

```
图(1)     图（1）    图1：说明    图[1]     图【1】    第〇章
Ａ（全角）  图·1      图—1        图+1      图%1      图@1
图#1      图*1       图♀1         图°1      …（共 22 个，仅 1 个被改写）
```

真正被改写的只有两类：**空白字符**（→ `_`）与**超过 40 字符**（→ 截断）。
所以 `validate_bookmark_name()` 的字符检查是**黑名单**——`ST_String` 实际禁止的控制字符
（U+0000–U+001F 与 U+007F），其余一律放行。

实测：

```python
validate_bookmark_name("图(1)")        # []
validate_bookmark_name("图（1）")      # []
validate_bookmark_name("第〇章")        # []
validate_bookmark_name("_Ref图 1")     # ["name contains whitespace (U+0020); ..."]
validate_bookmark_name("x" * 41)      # ["name is 41 characters, over the 40 limit; ..."]
validate_bookmark_name("图\x07")       # ["name contains control characters, ... U+0007"]
```

## 6. 幂等与安全：只删自己写过的东西

`toc 1` 是 Word 的内置样式名，**不是**所有权证明。用户从 Word 里"插入目录"生成的行
同样叫 `toc 1`；`drop_existing_toc()` 若按样式名删，就会静默抹掉用户自己的目录。
所以每个条目段落都带一个 `_DocxTocEntry…` 书签，**只删带标记的段落**。

实测（`claims.py`）：

```
1. 重复 build 三次 -> [(5, 11, 5), (5, 11, 5), (5, 11, 5)]  条数/段数/条目数三轮完全一致
2. 默认：用户自己的 toc 1 行还在 -> True | dropped_entries: 0 | ours: 5 | 段数 6 -> 12
   adopt_unmarked=True：接管 -> True | adopted: 1 | dropped_entries: 1
```

- 重复调 `build_static_toc` **不累积**：先删后建，每次都是同一批段落。
- `adopt_unmarked=False`（默认）：只删自己的。用户的行原样留着。
- `adopt_unmarked=True`：把无标记的 `toc…` 段落也当自己的删掉。这是"接管旧目录"的逃生口
  （旧版本本模块生成的、或用户明确要求替换 Word 生成的目录时用）。
  **删用户的目录不可恢复，所以默认关闭。**

`insert_at` 的语义是"**旧目录删完之后**，在 `doc.paragraphs` 的第几个段落之前插入"。
所以同一个 `insert_at` 值可以反复用；`insert_at=None` 时重建落在旧目录原来的位置
（`drop_existing_toc` 会返回 `anchor`＝被删第一条的前一个段落），不会跑到文档末尾。

**`--insert-at` 越界或负数一律 exit `2`、零写盘、stdout 为空**，合法区间 `[0, N-1]`
（`N` 是当前 body 段落数；空文档是唯一特例，只有 `0` 合法，代表追加）。

重建目录时这道检查要做**两遍**，因为删掉旧目录会让 body 变短：第一遍按调用方给的
文档判，第二遍按"旧目录已删"的 body 再判一次。第一遍合法、第二遍落空的索引会在
第二遍被拒，报错里写明是哪一遍没通过：

```
error: --insert-at 5 is out of range after the previous TOC was removed: legal range is now [0, 2], current N=3 (the index was legal against the body before removal; pass an index within the shortened body, or omit --insert-at to land where the old TOC was)
```

这里**刻意不退回"追加到末尾"**：那会把目录放到调用方从没要求的位置，而文档照样 exit `0`，
是最难发现的一种错位。要么给一个落在缩短后 body 里的索引，要么干脆不传 `--insert-at`。

> 一个实测到的边界：如果旧目录原本在文档最前面（`anchor is None`），
> `_body_insert_reference()` 会退化成"插到 body 第一个非 `sectPr` 子元素之前"，
> 结果落在原位——行为正确，但 `inserted_at` 可能比预期小 1，读返回值确认即可。

## 7. 多节页码：`monotonic` 已分节

前置 `i, ii` 接正文 `1, 2`，整串读下来是往回走的（2 → 1）。天真的单调检查会把一份**正确**的
学位论文目录判成坏的。`toc_audit.monotonic` 的做法是把序列**按计数器族切段**再逐段比较，
跨格式切换导致的"倒退"不算错。

`toc_audit` 的键（`probe_xml()["toc_audit"]`，顶层报告里也有同一份）：

| 键 | 含义 |
|---|---|
| `entries` | 带**级别**的 TOC 行数（样式名匹配 `^toc\s*\d`，`toc 1` / `TOC1` 都算） |
| `entries_toc_style` | 所有 TOC 系段落数（含无级别的 `目录…` / `contents…`） |
| `styles` | 用到的样式名，排序 |
| `levels` | 用到的级别，排序 |
| `page_tokens` | 每个条目尾部读到的页码 token（截到 `--limit-list`） |
| `page_tokens_total` | token 总数（未截断） |
| `page_parse_failures` | 尾部读不出页码的条目（`[{index, style, text}]`） |
| `monotonic` | **分族比较**后的结论：有没有真正倒退 |
| `monotonic_verdict` | 同一个结论的字符串形：`"true"` / `"false"` / `"unknown"` |
| `monotonic_across_sections` | 整篇不加切分的朴素结论（保留给需要它的调用方） |
| `monotonic_partial` | 有 token 读不出数字（`unparsed_tokens > 0`）时为 `True` |
| `unparsed_tokens` | 读不出数字的 token 条数 |
| `sections` | **TOC 段落自身落在**哪几节 |
| `source` | `"toc"`＝文档真的有 TOC 段落；`"computed"`＝没有（结构陈述，不是测量） |

读法：**先看 `monotonic_verdict`**。`"unknown"` 表示有 token 读不出来，此时 `monotonic`
仍是 `True` 但那是"读了的部分没倒退"，不是"目录没问题"。

实测（`semantics.py`，同一函数、不同 token 序列）：

```
decimal 1,5,3                    tokens=['1','5','3']        monotonic=False across=False
decimal 1,1,2                    tokens=['1','1','2']        monotonic=True  across=True
roman i,ii,iii                   tokens=['i','ii','iii']     monotonic=True  across=True
roman iii,i                      tokens=['iii','i']          monotonic=False across=False
front matter i,ii then body 1,2   tokens=['i','ii','1','2']   monotonic=True  across=False
body then roman (reverse order)   tokens=['1','2','i','ii']   monotonic=False across=False
letters a,b,c                    tokens=['a','b','c']        monotonic=True  across=True  partial=True
letters c,b,a                    tokens=['c','b','a']        monotonic=False across=True partial=True
```

第二行的 `across=False, monotonic=True` 就是分节的意义：整篇看是倒退，分族看是对的。
`a,b,c` 的 `partial=True` 说明单字母无法判定族（`c` 是罗马 100 还是字母 3），
所以结论降级为 `unknown`——宁可说"判不了"，也不给假绿灯。

### 7.1 大小写本身是两个计数器

Word 的 `upperLetter` / `lowerLetter`（以及 `upperRoman` / `lowerRoman`）是**四个独立计数器**，
所以 `A,B,C` 之后接 `a,b,c` 不是"页码倒退"，而是换了一个计数器重新从 1 数。

但大小写换族不像切节那样有 `w:start` 可查，只能从 token 形态推断，于是判据是
**"同一对相邻 token 上同时出现下坠和大小写翻转"**——两个信号重合才认定是换族重起：

```
A,B,a,b          → monotonic=False, verdict="unknown"   （看得见下坠，但可归因于换族）
I,II,III,i,ii    → monotonic=False, verdict="unknown"
A,B,C,a,b,c      → monotonic=False, verdict="unknown"   （下坠跨 run 边界，仍能配对）
M,D,C            → monotonic=False, verdict="false"     （纯下坠、无翻转：真倒退）
c,a,b            → monotonic=False, verdict="false"     （纯下坠、同族：真倒退）
A,a              → monotonic=True,  verdict="true"      （无下坠：同一计数器的重复读法）
A,B,E,F          → monotonic=True,  verdict="true"      （翻转但无下坠：不降级）
```

`monotonic` 仍如实报 `False`（"我确实看到了下坠"），`verdict` 才降 `unknown`（"但说不清是不是错"）。
这与 `w:start` 造成的合法重启是同一套语义：**看得见，但不判**。

## 8. 完整可跑的例子

三份文件都在 `%TEMP%` 之下，不写死任何本机路径。

### 8.1 spec（先生成正文，再插目录）

`paper_spec.json` —— 封面 / 前置（小写罗马）/ 正文（阿拉伯，重新从 1 开始）：

```json
{
  "page": {"width_mm": 210, "height_mm": 297,
           "margins_mm": {"top": 25, "bottom": 25, "left": 20, "right": 20}},
  "styles": [
    {"name": "卷标题", "base": "Normal", "font": "黑体", "size_pt": 16,
     "bold": true, "outline_level": 1},
    {"name": "节标题", "base": "Normal", "font": "黑体", "size_pt": 14,
     "bold": true, "outline_level": 2},
    {"name": "条标题", "base": "Normal", "font": "楷体", "size_pt": 12,
     "bold": true, "outline_level": 3}
  ],
  "blocks": [
    {"type": "paragraph", "text": "封面占位"},
    {"type": "section_break", "start_type": "new_page",
     "footer_page_numbers": {"fmt": "lowerRoman", "start": 1}},
    {"type": "paragraph", "text": "摘要占位"},
    {"type": "page_break"},
    {"type": "section_break", "start_type": "new_page",
     "footer_page_numbers": {"fmt": "decimal", "start": 1}},
    {"type": "paragraph", "text": "第一章　静电场", "style": "卷标题"},
    {"type": "paragraph", "text": "电场是电荷周围存在的一种特殊物质。"},
    {"type": "paragraph", "text": "1.1　库仑定律", "style": "节标题"},
    {"type": "paragraph", "text": "真空中两个静止点电荷之间的作用力。"},
    {"type": "paragraph", "text": "1.1.1　点电荷模型", "style": "条标题"},
    {"type": "paragraph", "text": "把带电体抽象为一个几何点。"},
    {"type": "paragraph", "text": "第二章　静电场中的导体", "style": "卷标题"},
    {"type": "paragraph", "text": "导体内部场强为零。"}
  ]
}
```

### 8.2 两趟写法（推荐：第一趟占位，第二趟填页码）

```python
"""toc_build.py — 目录两趟法。用法: python toc_build.py <skill-root> <docx>"""
import json, subprocess, sys
from pathlib import Path

skill = Path(sys.argv[1]).resolve()
target = Path(sys.argv[2]).resolve()
sys.path.insert(0, str(skill / "scripts"))

from docx import Document
import docx_create as dc
import docx_read
import docx_toc as toc

doc = Document(str(target))
dc.ensure_toc_styles(doc, levels=(1, 2), indent_chars=(0, 2))   # 外观

# 第一趟：只要条目，页码留空槽。这一步会把段数定下来。
first = toc.build_static_toc(doc, levels=(1, 2), title="目录", insert_at=1)
doc.save(target)
print("pass 1:", {k: first[k] for k in
      ("entries", "levels", "styles", "tab_pos_mm", "tab_pos_source",
       "section_index", "inserted_at")})

# 哪些段落是标题，用 Word 自己的 1-based 序号
doc = Document(str(target))
word_indices = [i + 1 for i, p in enumerate(doc.paragraphs)
                if docx_read.heading_level(p)]

# 问 Word：每个标题落在哪一页、那一页印的是什么
probe = subprocess.run(
    [sys.executable, str(skill / "tools" / "word_accept.py"), str(target),
     "--page-probe", ",".join(map(str, word_indices)), "--quiet"],
    check=True, capture_output=True, text=True, encoding="utf-8")
report = json.loads(probe.stdout)
assert report["paragraph_pages_source"] == "word", report["paragraph_pages_source"]
pages = [r["displayed_text"] for r in report["paragraph_pages"]]

# 第二趟：同一个 insert_at，把页码写进去（旧目录先被删掉，不会累积）
doc = Document(str(target))
second = toc.build_static_toc(doc, levels=(1, 2), pages=pages, title="目录",
                              insert_at=1)
doc.save(target)
print("pass 2:", {k: second[k] for k in ("entries", "dropped_entries",
      "dropped_titles", "tab_pos_mm", "tab_pos_source")})

for para in Document(str(target)).paragraphs:
    if toc.is_toc_paragraph(para) or toc.is_toc_title(para):
        print("  TOC | %-12s | %s" % (para.style.name,
                                      para.text.replace("\t", "<TAB>")))
```

### 8.3 命令行复现（PowerShell，可移植）

下面整段已于 2026-11-11 在仓库根端到端复跑（本机 Word 16.0），逐条退出码 `0`。

```powershell
$work = Join-Path $env:TEMP 'toc-demo'
New-Item -ItemType Directory -Force -Path $work | Out-Null

# 仓库根（含 scripts/ tools/）；<skill-root> 换成你的检出目录
Set-Location <skill-root>

# paper_spec.json / toc_build.py 放到 $work 下（内容见 §8.1、§8.2）

python scripts/docx_create.py "$work\paper_spec.json" "$work\paper.docx"

# 两趟：第一趟占位 + 问 Word 页码 + 第二趟回填
python "$work\toc_build.py" (Get-Location).Path "$work\paper.docx"

# 只想看页码真相（I,J,... 是 Word 的 1-based 段落序号）
python tools/word_accept.py "$work\paper.docx" --page-probe 8,10,12,14 --quiet

# 目录质量自检（--word off，不需要 Office）
python tools/word_accept.py "$work\paper.docx" --word off --json "$work\audit.json" --quiet |
  Out-Null
python -c "import json;print(json.dumps(json.load(open(r'$work\audit.json',encoding='utf-8'))['toc_audit'],ensure_ascii=False,indent=2))"
```

`docx_create.py` 打印的是一行 JSON（`{"ok": true, "output": ..., "blocks": 13, ...}`），
`word_accept.py` 的 stdout 也**恰好一个 JSON 文档**，人读摘要走 stderr。
所以上面所有 `|` 管道都能安全地喂给 `python -c`。

实测输出（上面 §8.1 的 spec + §8.2 的 `toc_build.py`，本机 Word 16.0，逐字复制）：

```
{"ok": true, "output": "...\\paper.docx", "blocks": 13, "sections": 3, "crossrefs": 0}
pass 1: {"entries": 3, "levels": [1, 2], "styles": ["toc 1", "toc 2"], "tab_pos_mm": 170.0, "tab_pos_source": "content_width_mm", "section_index": 0, "inserted_at": 1}
page source: word
  para 8   sec=3 physical=3   printed='1'  fmt=decimal
  para 10  sec=3 physical=3   printed='1'  fmt=decimal
  para 12  sec=3 physical=3   printed='1'  fmt=decimal
  para 14  sec=3 physical=3   printed='1'  fmt=decimal
pass 2: {"entries": 3, "dropped_entries": 3, "dropped_titles": 1, "tab_pos_mm": 170.0, "tab_pos_source": "content_width_mm"}
  TOC | TOC Heading  | 目录
  TOC | toc 1        | 第一章　静电场<TAB>1
  TOC | toc 2        | 1.1　库仑定律<TAB>1
  TOC | toc 1        | 第二章　静电场中的导体<TAB>1
```

几点读法：

- `entries` 是 3 而不是 4，因为 `levels=(1, 2)` 把 `条标题`（第 3 级）排除了。
- `section_index: 0` 说的是**制表位按第 1 节的版心算**——目录插在 `insert_at=1`，
  也就是封面那一节里。§9 会讲这个推算的边界。
- 三行页码都是 `1`，是因为这份样本每节只有一页，不是缺陷。

更完整的对照样本（前置罗马 + 正文阿拉伯，页码真的不同）：

```
pass 1   {"entries": 4, "levels": [1, 2], "styles": ["toc 1", "toc 2"],
          "tab_pos_mm": 170.0, "tab_pos_source": "content_width_mm",
          "section_index": 2, "inserted_at": 1}
source   word
  para 7   sec=2 physical=2   shown='i' fmt=lowerRoman
  para 10  sec=3 physical=3   shown='1' fmt=decimal
  para 13  sec=3 physical=3   shown='1' fmt=decimal
  para 16  sec=3 physical=4   shown='2' fmt=decimal
pass 2   {"entries": 4, "dropped_entries": 4, "dropped_titles": 1, ...}
  TOC | TOC Heading  | 目录
  TOC | toc 1        | 摘要<TAB>i
  TOC | toc 1        | 第一章　静电场<TAB>1
  TOC | toc 2        | 1.1　库仑定律<TAB>1
  TOC | toc 1        | 第二章　静电场中的导体<TAB>2
```

`摘要` 那行写的是 `i` 而不是 `2`——这就是 §3 那张表的意思。

### 8.4 CLI 的 `--pages` 必须一条目一值（Python API 反而宽容）

`docx_toc.py --pages` 是**按序**落位的：第 n 个值写进第 n 个条目。所以个数不等就是打错，
不是风格问题——多出来的被丢弃、缺的留成空槽，两者都静默，而交付文件里夹一个空页码与
"压根没填页码"看起来一模一样。因此 CLI 在**动文档之前**先数条目：

```powershell
python scripts/docx_toc.py paper.docx -o out.docx --title 目录 --levels 1,2 --pages 1,1,1
```

```
{"entries": 3, "levels": [1, 2], "styles": ["toc 1", "toc 2"], "tab_pos_mm": 170.0, ...}
```

给错个数（这里是同一份文档，默认 `--levels` 为 1,2,3 → 4 个条目）：

```powershell
python scripts/docx_toc.py paper.docx -o out.docx --title 目录 --pages 1,3,7
```

```
error: --pages has 3 value(s) but the TOC has 4 entries; give exactly one page number
per entry, or omit --pages to leave the slots empty
```

退出码 `2`，`out.docx` **不生成**。要点：数的是**当前 `--levels` 选中的条目数**，不是
文档里所有标题数——上面 `--levels 1,2` 就是 3 条、默认 `1,2,3` 就是 4 条。Python 侧
`build_static_toc(pages=[...])` 故意保持宽容（缺的留空槽、多的不用），因为"故意只填一部分"
是调用方的合法意图，与 shell 用户少打一个逗号不是一回事。

## 9. 右制表位：`tab_pos_mm` 与 `tab_pos_source`

页码靠右的对齐点必须正好落在**版心右边缘**，否则每一条点线都会偏。位置按优先级取：

| `tab_pos_source` | 来源 | 何时出现 |
|---|---|---|
| `tab_pos_mm` | 调用方显式给的值 | 你量过模板 |
| `content_width_mm` | `docx_create.content_width_mm(section)`＝页宽 − 左 − 右 | 默认，最常用 |
| `fallback` | `DEFAULT_TAB_POS_MM = 160.0`（A4 + 30/20 边距，房式） | 上面两条都拿不到 |

实测（同一 `build_static_toc` 调用，不同参数）：

```
default (no hint)                      -> tab_pos_mm=160.0  source=content_width_mm  section_index=0
section_index=0 (160 mm 版心)          -> tab_pos_mm=160.0  source=content_width_mm  section_index=0
section_index=1 (170 mm 版心)          -> tab_pos_mm=170.0  source=content_width_mm  section_index=1
section_index=-1 (最后一节)            -> tab_pos_mm=170.0  source=content_width_mm  section_index=1
explicit tab_pos_mm=155                -> tab_pos_mm=155.0  source=tab_pos_mm        section_index=0
```

math 卷是 A4 20/20 → 170 mm；其余五卷 30/20 → 160 mm。**别把任何一个写死**，
这正是 `content_width_mm()` 存在的理由。

### `insert_at` 推节有偏差，落进多节文档就显式给 `section_index`

`_target_section()` 想用"`insert_at` 段落属于哪一节"来定版心宽度。实测两节文档
（第 1 节 160 mm，第 2 节 170 mm，段落 `[封面, '', 正文]`，段落 1 带 `sectPr`）：

```
insert_at=1 -> 160.0  section_index 0    正确（段落 1 是第 1 节最后一段）
insert_at=2 -> 160.0  section_index 0    错误（段落 2 是第 2 节正文，应为 170 mm）
explicit section_index=1 -> 170.0        正确
```

结论：**目录落在第一节以外时，显式传 `section_index`**。目录一般都在文档最前面，
所以日常不受影响；但 `tab_pos_source` / `section_index` 两个返回值要读一下确认。

## 10. 数字格式对照（`docx_section.cached_page`）

`w:pgNumType/@fmt` 决定"页码长什么样"。下表第 2 列是代码，第 3 列是本机 Word 16.0 实测
（页脚 PAGE 域 → `Fields.Update()` → 读回 `Result.Text`，page 1..3）：

| `fmt` | `cached_page(fmt, 1..3)` | Word 16.0 实测 | 一致 |
|---|---|---|---|
| `decimal` | `1, 2, 3` | `1, 2, 3` | ✅ |
| `lowerRoman` | `i, ii, iii` | `i, ii, iii` | ✅ |
| `upperRoman` | `I, II, III` | `I, II, III` | ✅ |
| `lowerLetter` | `a, b, c` | `a, b, c` | ✅ |
| `upperLetter` | `A, B, C` | `A, B, C` | ✅ |
| `decimalZero` | `01, 02, 03` | `01, 02, 03` | ✅ |
| `none` | `` `` （空串） | 页脚空串；正文里的 PAGE 域会显示 `错误!不能以指定格式显示编号。` | ✅（见下） |
| `decimalFullWidth` | `１, ２, ３` | `１, ２, ３` | ✅ |
| `chineseCounting` | `一, 二, 三` | `一, 二, 三` | ✅ |
| `ideographDigital` | `一, 二, 三` | `一, 二, 三` | ✅ |
| `numberInDash` | `- 1 -`, `- 2 -`, `- 3 -` | 同 | ✅ |
| `decimalEnclosedCircle` | `①, ②, ③` | 同 | ✅ |
| `ordinalText` | `first, second, third` | 同 | ✅ |
| `hex` | `1..9, 10, 11` | `1..9, A, B` | ❌ 从第 10 页起不一致 |

两点要注意：

- **`none`**：页脚里空（没有页码），这正是"这一节不印页码"的效果。但正文里放一个
  `PAGE` 域会显示 Word 的错误文本——那是域自身的限制，不是 `cached_page` 的错。
- **`hex`**：`cached_page("hex", 10)` 返回 `"10"`，Word 印 `A`。第 1–9 页碰巧一致，
  第 10 页起就错。**要 hex 页码的文档别依赖这个缓存值**（本条已报给实现方）。

## 11. 排错清单

| 症状 | 先查 |
|---|---|
| 静态目录一条都没有 | 样式有没有 `outline_level`？`collect_outline(doc)` 空不空？ |
| TOC 域 F9 之后写"未找到目录项。" | 同上；`w:outlineLvl` 是不是写在样式上而没被 Word 丢弃 |
| 目录页码是 `1,2,3` 但纸上是 `i,ii,iii` | 走 CLI，或者自己给 `probe_word(section_formats=...)` |
| 页码看着像物理页序 | 用了 `physical` 而不是 `displayed_text` |
| 点线没对齐 / 跑出版心 | `tab_pos_source` 是什么？多节文档给了 `section_index` 吗？ |
| 重复生成后出现了两份目录 | 第二份是不是用户自己用 Word 插入的？`adopt_unmarked` 默认不会删它 |
| `toc_audit.monotonic` 是 `True` 但感觉不对 | 看 `monotonic_verdict`：`unknown` 表示有 token 读不出来 |
| 样式名在 XML 里是数字（`w:pStyle w:val="9"`） | 那是 `styleId` 不是名字；`probe_xml` 已解析成 `style_usage` |
