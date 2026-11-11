# 表格几何技术参考

## OOXML 怎么解析列宽

Word/WPS 实测（2026-10-27）：

- `tblLayout` 缺失（autofit）→ 把每格 `<w:tcW>` 当**宽度提示**，列宽按 tcW 比例分配。
- `tblLayout w:type="fixed"` → **只认** `<w:tblGrid><w:gridCol>`，**忽略 tcW**。grid 若是默认
  等分，不管 tcW 写什么每列都一样宽。

第二条是陷阱本体：**python-docx 的 `table.autofit = False` 写的正是 `tblLayout=fixed`**，所以
只设 `cell.width` 会静默得到等宽列。三处必须同时写且自洽：

```xml
<w:tbl><w:tblPr><w:tblW w:type="dxa" w:w="TOTAL"/>
          <w:tblLayout w:type="fixed"/>
       <w:tblGrid><w:gridCol w:w="W1"/><w:gridCol w:w="W2"/>...
       <w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="W1"/>
```

单位一律 dxa：**1 cm = 566.93 dxa**，`cm_to_dxa(cm) = int(round(cm / 2.54 * 1440))`。
`tblW` = 各列之和；gridCol = 逐列宽（fixed 下才是渲染依据）；tcW = 该格跨列之和
（`gridSpan > 1` 要累加）。

## 三处宽度必须一致

上面那段骨架的三处不是「都写上就完事」，而是必须**彼此一致**：无论渲染器把哪一处当权威，另外两处都得跟它对齐，否则同一个文件的表现会随改动的先后顺序变——先动 `tblLayout` 再动 `tcW`，与反过来做，出来的列宽能不一样。原文这条规则是：

> Keep all three consistent. Whichever element is authoritative, the other two must agree with it, or you get behaviour that changes depending on which toggle happened last.

落到操作上就是三句：`tblW` = `gridCol` 之和；`gridCol` 逐列等于该列应有宽度；`tcW` 等于该格跨列宽度之和。只改其中一处，另外两处立刻成了新的错配源——后面「常见症状与修法」一节列的四种表现，全都是这个原因。

## 四个子命令

`action` 只接受 `report` / `set` / `sync` / `auto`。`unpacked_dir` 下必须有
`word/document.xml`。表格序号 1-based；不给 `--table` 也不给 `--all` 时作用于全部。

| 模式 | 命令 | 行为 |
|---|---|---|
| `report` | `table_width.py report mydoc/` | 只看不改，打 tblW/tblLayout/gridCol |
| `set` | `set mydoc/ --table 1 --widths 2.4,4.2,2.4,5.6` | 显式给 cm，列数不足补末列 |
| `sync` | `sync mydoc/` | 把第一行 tcW 抄进 grid 再钉死 |
| `auto` | `auto mydoc/ --content-width 14.6` | 文本宽按列均分（默认 14.6 cm） |

**`--widths` 个数与列数不等时，两个方向的处理并不对称**（2026-11-11 实测，2 列表）：

- **多于列数 → exit 2，零写入**：`set --widths 2.4,4.2,2.4,5.6`（4 个值给 2 列）报
  `error: table 1: --widths has 4 values but the table has 2 columns; refusing to drop the extra 2`，
  退出码 `2`，`word/document.xml` 的 SHA256 不变。丢多余的宽度等于静默改写调用方的意图，所以拒收。
- **少于列数 → 补末列，但 stderr 必须说明**：`set --widths 3.0`（1 个值给 2 列）退出 `0`、
  两列都写成 3.00 cm，同时 stderr 打
  `warning: table 1: --widths has 1 values but the table has 2 columns; padded 1 column(s) with 3.00 cm`。
  补末列本身是刻意保留的便利（模板常只量第一列），但**不申报的补全就是在发明调用方没给的宽度**，
  所以补齐一定留痕。

`report` 还会做 tcW vs gridCol 比对（容差 `TOL = 20` dxa ≈0.35 mm），超过打
`! col N: tcW=… != gridCol=…`；若同时 `fixed` 再补 `tblLayout=fixed ignores tcW`，提示跑 `sync`。

**report 先跑**：它把 grid、各格 `tcW`、以及两者是否打架一次打全。不要一上手就 `set`，因为该修哪一步取决于
眼前到底是哪种错配——grid 没跟上 tcW 就是 `sync`，宽度本身就不对才轮到 `set`，盲目先 `set` 会把
本来正确的比例一起改掉。

`set` 的 `--widths` 须全解析成 float，否则报 `--widths must be comma-separated numbers in cm`
+ exit 1；缺 `--widths` 报 `'set' requires --widths`（同样 exit 1）。

**退出码实测矩阵**（2026-11-11，全部在仓储根跑）：

| 场景 | 退出码 | stderr |
|---|---|---|
| `report` 一份真没有表格的文档 | `0` | 无（stdout `no tables found`）—— 读到一个空结果是真答案 |
| `set` / `sync` / `auto` 打一份没有表格的文档 | `1` | `no tables found; nothing written`（一个宽度都没写，`0` 会让调用方以为钉住了） |
| `--widths` 多于列数 | `2` | 见上，且零写入 |
| `--widths` 不是数字 / 缺 `--widths` | `1` | `--widths must be comma-separated numbers in cm` / `'set' requires --widths` |

> 一处口径不一致，留待实现方：`scripts/table_width.py` 的 docstring 把「`word/document.xml`
> 缺失」归到 exit `2`，实测该分支 `return 1`（`error: …\word\document.xml not found`）。
> 本文按实测写；`scripts/` 不在本轮文档修订范围内。

`sync` 的 `_read_tcw()` 从**第一行**读 `<w:tcW>`（`w:type` 须为 `None`/`dxa`，`gridSpan > 1`
按 `v / span` 均摊），拿不到退回 `_read_grid()`；两者都空则 `skip table N` 并跳过。

### 写盘（三模式共用 `_apply_widths`）

`_ensure_grid` 保证 `tblGrid` 行列对齐并逐 gridCol 写 `w:w`；`_set_tbl_width` 写 `tblW` =
各列之和；`_set_fixed_layout` 写 `tblLayout type="fixed"`；最后遍历 `w:tr`/`w:tc` 按
`gridSpan` 累加写 `tcW`。行高不动。

## 常见症状与修法

| 症状 | 成因 | 修法 |
|---|---|---|
| 每列一样宽 | `tblLayout=fixed` 但 grid 还是默认等分 | `set --widths` 显式给宽度 |
| 编辑某一格后列宽整体跳变 | autofit 模式在按 `tcW` 重算 | 先钉 `fixed`，再跑 `sync` |
| 第一列对、其余全错 | grid 与 `tcW` 不一致 | 先 `report` 看差在哪，再 `sync` |
| 总宽超出页面 | `tblW` 与 grid 之和不等 | `set`（会重算 `tblW`） |

四行的共同根因都是「三处宽度只有一处是对的」。`report` 打出来的 `! col N: …` 就是定位工具。

## w:zoom 的 schema 坑

`docx_create.py` / `docx_template.py` 生成的文档，python-docx 写的是
`<w:zoom w:val="bestFit"/>`。CT_Zoom 把 `w:percent` 声明为 `use="required"`，严格 XSD 直接拒：
`Element '{...}zoom': The attribute '{...}percent' is required but missing.`

```
python scripts/fix_zoom.py file.docx                # 就地修复
python scripts/fix_zoom.py file.docx -o out.docx    # 写到别处
python scripts/fix_zoom.py file.docx --percent 120  # 非默认缩放
```

正则必须是这条：

```python
ZOOM_RE = re.compile(r"<w:zoom\b([^>]*?)(/?)>", re.S)
PERCENT_ATTR_RE = re.compile(r'\bw:percent\s*=\s*"[^"]*"')
```

裸正则 `<w:zoom\s*/>` **一个带属性的形式都匹配不上**——脚本"成功"了但文档依旧非法。这是最
隐蔽的失败模式：没有异常、没有非零退出码，只有 Word 打不开。

另三条：已有 `w:percent` 的不能再加（重复属性让 XML 非法），命中即原样返回；zip 重写必须
**复用原 `ZipInfo`**，条目顺序才一致；没有 `word/settings.xml` 的包合法（报 `nothing to do` +
exit 0）；幂等，第二次跑报 `zoom element already valid`。纯函数入口
`fix_settings_xml(xml, percent) -> tuple[str, int]`（`tests/` 直接 import）。

## 解包自检

```python
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
with zipfile.ZipFile("out.docx") as z:
    root = ET.fromstring(z.read("word/document.xml"))

for i, tbl in enumerate(root.iter(W + "tbl"), 1):
    grid = [int(g.get(W + "w")) for g in tbl.iter(W + "gridCol")]
    total = int(tbl.find(f"{W}tblPr/{W}tblW").get(W + "w"))
    assert sum(grid) == total, f"table {i}: gridCol {sum(grid)} != tblW {total}"
    layout = tbl.find(f"{W}tblPr/{W}tblLayout")
    assert layout.get(W + "type") == "fixed"
    for tr in tbl.findall(W + "tr"):
        row = sum(int(tc.find(f"{W}tcPr/{W}tcW").get(W + "w")) for tc in tr.findall(W + "tc"))
        assert row == total, f"table {i}: 该行 tcW {row} != tblW {total}"
```

失败基本就三种：gridCol 没写、tblW 是 `auto` 不是 `dxa`、`gridSpan` 格的 tcW 没累加。
再跑 `python scripts/docx_validate.py out.docx`——它是**健康检查不是 XSD 校验**，`w:zoom` 缺
`w:percent` 这类问题抓不到，那件事归 `scripts/fix_zoom.py`。
