# `tools/` — 验收工具（Wave 2 全体 agent 共用）

`word_accept.py` 是本项目**唯一**的 Word 验收手段。所有"可打开性 / 页数 / 节数 /
书签 / 域结果 / 编号渲染 / 超链接可点击性 / PDF / 逐页 PNG"都由它一次采齐，输出一份
稳定 JSON。**不要自己再写 Word COM 脚本**——每多一份，证据的口径就多一种。

> 只读工具：它从不修改被检文件（Word 里 `ReadOnly=True`，`Close(False)` 永不保存）。

---

## 1. CLI 契约（冻结，不要改形状）

```
python tools/word_accept.py <in.docx> [--pdf OUT.pdf] [--png DIR] [--json OUT.json]
        [--expect-sections N] [--expect-bookmarks N] [--expect-params N]
```

| 参数 | 说明 |
|---|---|
| `<in.docx>` | 被检文件（必填） |
| `--pdf OUT.pdf` | 用 Word 导出 PDF（`ExportAsFixedFormat(..., 17)`，失败自动退 `SaveAs2`） |
| `--png DIR` | 把 PDF 每一页渲染成 `DIR/page-NNN.png`（需同时给 `--pdf`，PyMuPDF，`--dpi` 默认 110） |
| `--json OUT.json` | 把同一份 JSON 落盘（UTF-8 无 BOM）；stdout 仍然打印 |
| `--expect-sections N` | 断言 `sections == N` |
| `--expect-bookmarks N` | 断言 `bookmarks == N` |
| `--expect-params N` | 断言 `paragraphs == N`（**段数**；别名 `--expect-paragraphs`） |
| `--word auto\|on\|off` | 默认 `auto`：有 Word 就用。`off`＝只跑 XML 层。`on`＝要求 Word，拿不到就报错 |
| `--numbering-scan numpr\|all` | 默认 `numpr`（走 `Document.ListParagraphs`，快）；`all` 是逐段慢扫的逃生口 |
| `--no-update-fields` | 跳过 `Fields.Update()` |
| `--limit-list N` | 所有列表型字段的条数上限（默认 200） |
| `--quiet` | 不要 stderr 摘要 |

额外可用的断言（契约之外的新增，不影响上面三个）：
`--expect-pages` / `--expect-tables` / `--expect-fields` / `--expect-hyperlinks` /
`--expect-inline-shapes` / `--expect-openable yes|no`。

**入口/出口约定**

* stdout = **恰好一个 JSON 文档**，没有别的字。
* 人读摘要（计数、编号渲染、链接、错误、警告）走 **stderr**。
* 退出码：**0** 成功或 `skipped`；**1** `ok:false`；**2** 用法错误。

```powershell
# 典型调用（Wave 2 照抄）
# $py = 你的 Python 3.10+ 解释器（需装 python-docx / pywin32 / PyMuPDF）
$py = (Get-Command python).Source
& $py tools/word_accept.py out/exam.docx --pdf out/exam.pdf --png out/exam_png `
     --json out/exam.json --expect-sections 3 --expect-bookmarks 151
```

---

## 2. 两层结构（这是稳定性的来源）

| 层 | 函数 | 依赖 | 用途 |
|---|---|---|---|
| XML | `probe_xml(path) -> dict` | 只用 stdlib（`zipfile` + `xml.etree`） | CI 可跑；结构 / 书签 / 域码 / `w:hyperlink` 包裹 / `numPr`+`lvlText` / TOC 页码 / 样式普查 |
| Word | `probe_word(...) -> dict` | `pywin32` + Office | 只有 Word 才知道的：页数、节几何、渲染后的编号串、超链接**目标**、域**结果**、PDF、PNG |

`probe_xml` **刻意不 import `scripts/` 下任何东西**（含 `_bootstrap`），
`tools/word_accept.py` 里自带一份 6 行的 `_fix_stdio`。这样即使 `scripts/` 正在被并行改写，
验收器照常可用。同理 `win32com` / `fitz` 都只在函数内延迟导入——
`tests/test_word_accept_xml.py` 用 `ast` 把这条钉死了。

`main()` 把两层合并。**顶层契约字段永远存在且已填值**：Word 没跑时回退到 XML 的值，
并靠 `sources` 告诉你每个数字来自哪一层（`"word"` / `"xml"` / `"toc"` / `"none"`）。
调用方永远不用为"键不存在"写兜底。

---

## 3. JSON 字段

```json
{
  "tool": "word_accept", "version": "1.0", "docx": "...",
  "ok": true, "openable": true, "skipped": false, "word_available": true,

  "pages": 22, "paragraphs": 287, "sections": 3, "bookmarks": 151,
  "inline_shapes": 10, "fields": 2, "hyperlinks": 2, "tables": 4,

  "bookmark_names": ["图_1"],
  "hyperlink_targets": [{"sub_address": "图_1", "text": "1",
                         "address": "", "type": 0,
                         "superscript": 9999999, "superscript_on": true}],
  "field_list": [{"code": "REF 图_1 \\h", "result": "1"}],
  "numbering_render": [{"index": 272, "index_xml": 246, "index_source": "word",
                        "list_string": "[1]", "list_type": 3, "ilvl": "0",
                        "num_id": "10", "lvl_text": "[%1]",
                        "rendered": true, "text_match": true, "text": "盛万兴..."}],
  "numbering_summary": {"xml_numpr": 16, "word_list_paragraphs": 16,
                        "rendered": 16, "scan": "numpr", "text_mismatch": 0},
  "xref_summary": {"hyperlinks": 2, "targets": 2, "internal_anchors": 2,
                   "external": 0, "superscript": 1, "normal_size": 1,
                   "clickable": true, "wrapper_count_xml": 2, "ref_fields": 2},
  "page_numbers": ["i", "ii", "1", "2"], "page_numbers_source": "toc",
  "pdf": "out/exam.pdf", "pngs": ["out/exam_png/page-001.png"],
  "pdf_pages": 22, "pdf_page_footers": ["", "1", "2"],
  "section_details": [{"index": 1, "page_width_mm": 210.01, "page_height_mm": 297.0,
                       "top_margin_mm": 24.99, "bottom_margin_mm": 24.99,
                       "left_margin_mm": 30.0, "right_margin_mm": 20.0,
                       "orientation": 0, "headers": 3, "footers": 3,
                       "last_page": 22}],
  "errors": [], "hard_errors": [], "warnings": [],
  "expectations": [{"name": "sections", "expected": 3, "actual": 1,
                    "pass": false, "source": "word"}],
  "sources": {"paragraphs": "word", "hyperlinks": "word", "pages": "word", ...},
  "xml": { ...probe_xml 全量... },
  "word": {...} | null
}
```

### 3.1 顶层字段的来源与含义

| 字段 | Word 在跑 | Word 没跑（`--word off` / 无 Office） |
|---|---|---|
| `pages` | `ComputeStatistics(2)` | `null`（XML 无法推算） |
| `paragraphs` | `Paragraphs.Count` | `xml.paragraphs_xml` |
| `sections` | `Sections.Count` | `xml.sections_xml`（`w:sectPr` 个数） |
| `bookmarks` | `Bookmarks.Count` | `xml.bookmarks_xml`（含 `_` 前缀内部书签） |
| `inline_shapes` | `InlineShapes.Count` | `xml.inline_drawings_xml` |
| `fields` | `Fields.Count` | `xml.fields_xml` |
| `hyperlinks` | `Hyperlinks.Count` | `xml.hyperlinks_xml`（`w:hyperlink` 元素数） |
| `tables` | `Tables.Count` | `xml.tables_xml` |
| `bookmark_names` | 非 `_` 开头的书签名 | 同左，来自 XML |
| `hyperlink_targets` | `Hyperlinks(i).SubAddress` / `Address` / `Font.Superscript` | `xml.hyperlink_anchors`（`superscript` 为 `null`） |
| `field_list` | `Fields(i).Code.Text` / `Result.Text`（`Fields.Update()` 之后） | XML 状态机解析出的域码与缓存结果 |
| `numbering_render` | 按**序数**与 `ListParagraphs` 配对，带 `list_string` | 只有 XML 侧信息，`list_string`/`rendered` 为 `null` |
| `page_numbers` | 优先 `PAGEREF` 域结果 | TOC 条目尾部页码（不需要 Word！） |
| `pdf` / `pngs` / `section_details` | 有 | `null` / `[]` |

### 3.2 `ok` / `openable` / `errors` / `hard_errors` 的语义

* `openable`：Word 打开了＝`true`；Word 打不开＝`false`；**没跑 Word**＝`null`（"不知道"，
  不是"不行"）。唯一的例外：文件连 zip 都不是 / 缺 `word/document.xml` /
  `document.xml` 不是合法 XML → 直接 `false`，而且**不会拿去喂 Word**（避免 Word 弹修复框）。
* `ok`：`hard_errors` 或断言被**确凿证伪**时为 `false`；`openable` 为 `true` 且无硬错时为
  `true`；否则 `null`。
* `errors`：给人看的全部问题，**包含"Word 不可用，已跳过"这类信息性条目**。
* `hard_errors`：只有它会把 `ok` 打成 `false`。判断"这文件到底有没有问题"看它。
* 断言**测不出来**（值为 `null`）时进 `warnings` 并标 `"note": "unverifiable without Word"`，
  **不算失败**——测不了的断言不是缺陷。

---

## 4. ★ 必须知道的五个硬事实（都是实测踩出来的）

### 4.1 `gencache.EnsureDispatch` 同一进程只能用一次

同一进程里第二次 `win32.gencache.EnsureDispatch("Word.Application")` 会炸：
先 `RPC 服务器不可用`，再 `can not automate the makepy process`（makepy 缓存被占）。

**本工具的解法**：用 `win32.Dispatch`，并且**整个进程只建一次 app**
（模块级 `_STATE["app"]` 缓存，`atexit` 注册 `release_app()`）。
`main()` 结束时统一 `release_app()` → `Quit()`。

> 你要是自己写脚本：别用 `EnsureDispatch`。要用就保证一个进程里只调一次，
> 且退出前 `Quit()`。想在一个进程里验多个文件，就 `import word_accept` 然后循环调
> `probe_word(...)` —— app 会复用，不会踩这个坑。

### 4.2 `Paragraphs(i)` 索引 ≠ XML 里的 `w:p` 序号

**Word 给每一行表格额外算一个段落**（行结束标记）。实测 `f5.docx`：
XML `w:p` = **261**，Word `Paragraphs.Count` = **287**，`w:tr` = **26**。
`src.docx` 同理。

所以**永远不要**把 XML 序号当 `Paragraphs(i)` 用。本工具的办法：

* `numbering_render[].index_xml` 是 XML 序号，`index` 是 Word 序号，两者都给，别混用；
* 更关键的是**根本不做索引换算**——`probe_xml` 找出带 `w:numPr` 的段（文档序），
  Word 侧用 `Document.ListParagraphs`（同为文档序），两边**按序数配对**，
  并用前 12 字符文本比对，不符就写进 `numbering_summary.text_mismatch` 并给警告。
* 例外：`paragraphs`（顶层）在 Word 侧是 `Paragraphs.Count`，在 XML 侧是 `w:p`
  总数，所以**两者天然不等**，差多少就是表格行数。要对数就用 `--word off` 固定用 XML 值。

### 4.3 编号只能看 `ListFormat.ListString`

`Range().Text` **不含自动编号文本**，用它判断"编号没渲染"必然误判。
真值来源是 `Paragraphs(i).Range.ListFormat.ListString`。

但逐段取 287 次 COM 往返很慢。**本工具只在两层都认为是列表的段上取值**：
XML 侧先筛出 `w:numPr` 的段（`f5.docx` 是 16 段），Word 侧用 `ListParagraphs`
（原生集合，也是 16 项）→ 只花 16 次往返，而不是 287 次。
`--numbering-scan all` 是逐段慢扫的逃生口，正常情况下不要用。

编号格式（`[%1]` 之类）由 `probe_xml` 从 `numbering.xml` 解出：
`numId → abstractNumId → lvl[ilvl].lvlText`，放在 `lvl_text` 与 `xml.number_formats`。

### 4.4 `\h` 不产生 `w:hyperlink`

REF 域的 `\h` 只是个开关，**不生成 `w:hyperlink` 元素**，超链接只在 Word **重算该域之后**
才在内存里存在。写了缓存结果 → Word 打开时不重算 → **点不动**。这就是"交叉引用没有生效"的根因。

正解（BRIEF-COMMON §7.2）：把整个 REF 域**包进真正的**
`<w:hyperlink w:anchor="书签名" w:history="1">`。

**可点击性的程序化判定（本项目第一条能自动判定的可点击性证据）**：

```python
xs = report["xref_summary"]
assert xs["clickable"] is True          # 有链接且每条都有 sub_address/address
assert xs["internal_anchors"] >= 1      # Word: Hyperlinks(i).SubAddress 非空
assert xs["wrapper_count_xml"] >= 1     # XML: w:hyperlink 元素数（CI 也能断言）
```

验收时**同时看 `hyperlinks`、`hyperlink_targets[*].sub_address`、`wrapper_count_xml`**。
两层应当一致：`Word.Hyperlinks.Count == xml.anchor_hyperlinks_xml`。
若 Word 的数是 XML 的整数倍，说明**包了两层 wrapper**（schema 不允许嵌套
`w:hyperlink`）——比如 `cmd_ref` 已经包过一次，调用方又包了一次。这时别改验收器，去查调用方。

正常大小 vs 上角标：`hyperlink_targets[*].superscript_on`（`Font.Superscript`
`0` = 正常，`9999999`/`-1` = 上角标）。主人要求两种都要有，所以两个都要报。

### 4.5 `w:tab` 是元素，不是字符

`<w:tab/>` 不产生字符。如果把段落的 `w:t` 直接拼起来，TOC 条目
`"摘要" + <w:tab/> + "i"` 会变成 `"摘要i"`，尾巴上的页码就再也分不出来。
`_para_text()` 显式把 `w:tab` 还原成 `\t`、`w:br`/`w:cr` 还原成 `\n`。

另：`<w:pStyle w:val="9">` 里是 **styleId**，不是样式名。`src.docx` 的
`9/12/6` 名字是 `toc 1/toc 2/toc 3`，`24/25/26/27/28` 是
`章标题/节标题/条标题/正文/图表名`。任何按**样式名**匹配的逻辑都必须先经
`styles.xml` 解析，否则看到的是数字——上游"中文样式召回 6.5%"就是这一类坑。
`probe_xml` 已解析好：`xml.style_usage`（含 `style_id` / `name` / `count`）、`xml.toc_styles`。

---

## 5. 各学科 agent 建议的断言写法

```python
import json, subprocess, sys
rep = json.loads(open("out/my.json", encoding="utf-8").read())

# 0) 通用：文件没坏、Word 能开
if rep["skipped"]:
    print("警告：没跑 Word，只能核对结构；", rep["errors"]); 
assert rep["hard_errors"] == [], rep["hard_errors"]
assert rep["ok"] is True

# 1) 结构对得上
assert rep["sections"] == 3
assert rep["bookmarks"] == 151
assert rep["tables"] == 4

# 2) 编号真的渲染了（不是"XML 里有 numPr"就算完）
ns = rep["numbering_summary"]
assert ns["xml_numpr"] == ns["word_list_paragraphs"] == ns["rendered"], ns
assert ns["text_mismatch"] == 0
assert [e["list_string"] for e in rep["numbering_render"][:3]] == ["[1]", "[2]", "[3]"]
assert rep["numbering_render"][0]["lvl_text"] == "[%1]"

# 3) 交叉引用真的能点
xs = rep["xref_summary"]
assert xs["clickable"] and xs["internal_anchors"] >= 1, xs
assert all(t["sub_address"] for t in rep["hyperlink_targets"])

# 4) 页码（不需要 Word 也能拿到 TOC 那份）
assert rep["page_numbers"][:4] == ["i", "ii", "1", "2"]
assert rep["page_numbers_source"] == "toc"

# 5) 版面（样例回归：210×297mm，上25 下25 左30 右20）
s = rep["section_details"][0]
assert abs(s["page_width_mm"] - 210) < 0.2 and abs(s["page_height_mm"] - 297) < 0.2
assert abs(s["left_margin_mm"] - 30) < 0.2 and abs(s["right_margin_mm"] - 20) < 0.2

# 6) 视觉证据：PDF + 逐页 PNG 已落盘
assert rep["pdf"] and len(rep["pngs"]) == rep["pdf_pages"]
```

**CI 上没 Word 怎么跑**：加 `--word off`，只断言 `rep["xml"][...]` 与 `rep["sources"]`。
`page_numbers`、`style_usage`、`hyperlink_anchors`、`numpr_paragraphs`、`field_list_xml`
这些都不需要 Word，覆盖了绝大多数"结构/编号/引用"回归。

---

## 6. 已知限制 / 没做的事

1. **`pages` 只有 Word 能给**。`--word off` 时是 `null`，不要拿它做断言。
2. **页码只覆盖 TOC 那份**：`page_numbers` = `PAGEREF` 域结果（有的话）否则 TOC 条目尾部
   页码。文档若有独立页码域（页脚 `PAGE`），只有 `section_details[*].footer_page_number`
   给每节的第一个页码值，不是逐页列表。逐页页脚原文在 `pdf_page_footers`（需 `--pdf`）。
3. `paragraphs` 在两层口径不同（§4.2），不要跨层比较。
4. **`numPr` 与 `ListParagraphs` 的序数配对**假定两者都是文档序且数量相同。
   数量不同时多余项会照实列出并带 `note`，`text_mismatch` 会给计数——**这是告警不是结论**。
5. **文本框内的段落**：XML 计入 `paragraphs_xml`，Word 是否同样计入未在本机样本上验证
   （样本 `textboxes = 0`）。有文本框时会写 `xml.xml_warnings`。
6. **超链接 `text` 取不到 `TextToDisplay` 时回退 `Range.Text`**；外部链接的显示文本可能仍为空。
7. **`--png` 需要 `--pdf`**；单独给 `--png` 不会凭空造 PDF。
8. **未实现**：逐页页码数组、页眉逐页文本、浮动形状（`Shapes`）明细（只报了计数）、
   批注/修订计数（应交给 `docx_comments` / `docx_revisions` 自己的 CLI）。
9. `--word on` 在拿不到 Word 时会让 `ok:false`（这是你要的），`auto` 则静默降级为 skipped。
10. **不要 `import word_accept` 后长期驻留**（除非你要复用那个 Word 进程）——收尾记得
    `release_app()`，否则会留一个不可见的 WINWORD.EXE。CLI 模式已自动处理。

---

## 7. 自检

```powershell
$py = (Get-Command python).Source
Set-Location skill-patched
& $py -m pytest tests/ -q -p no:cacheprovider          # 全绿
& $py -m pytest tests/test_word_accept_xml.py -q -p no:cacheprovider
& $py -c "import ast;ast.parse(open('tools/word_accept.py',encoding='utf-8').read(),feature_version=(3,10));print('AST310 OK')"
```

回归正例样本可用 `work/agents/accept/make_sample.py` 重新生成
（它是用本仓库自己的 `docx_crossref` / `docx_numbering` 造的，含编号 + 题注 +
两种大小的可点击引用 + 表格）。
