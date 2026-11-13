# 字体方案技术参考

> 本文覆盖 `scripts/fix_fonts.py`（`w:rPr` 两条铁律 + `--heading-threshold`）、它依赖的
> `docx_utils` 三组字体常量，以及它与 `math_omml.apply_math_font()` 的交界。`docx_create.py`
> 的**样式级** East Asian 槽绑定是同一件事的另一条入口，见下节。判据都是本机 Word/WPS 实测。

## 两个入口：`fix_fonts.py` 与 `docx_create.py` 的样式级绑定

把中文字体绑到 `eastAsia` 槽、把 `w:rFonts` 摆进 `w:rPr` 的合法位置——同一件事有**两个入口**，
别让它们分叉：

| | `fix_fonts.py`（本文主角） | `docx_create.py`（`add_styles` 一族） |
|---|---|---|
| 时机 | **事后修**既有 `.docx`：遍历 `styles.xml` / `document.xml` 就地改写 | **生成时建**样式：spec 里一次写定 |
| 绑定函数 | `set_fonts(rpr, ea)`（`scripts/fix_fonts.py:89`）——删旧 `w:rFonts` 后 `rpr.insert(0, rf)` | `_set_style_fonts(style, latin, east_asia)`（`scripts/docx_create.py:222`）——`rpr.insert_element_before(rf, "w:b", …)` |
| 怎么表达 | CLI 参数 `--heading-threshold` + 三层判定 | spec 键 `font` / `font_east_asia`（`SKILL.md` 的 `## Creating a document` 一节） |
| 测试锚点 | 本文「解包自检」的 `(dup, not_first, total)` 判据 | `tests/test_style_fonts.py`（5 条） |

> **怎么定位（2026-11-11 复核）**：`styles[]` 的键清单在 `SKILL.md` 的
> `## Creating a document` 一节里，原文是「styles[] entries take font, font_east_asia,
> size_pt, bold, italic, color, base, and outline_level」。**别引行号**：本文早前引的
> 「L112」是 Quick reference 表里的 `fix_fonts` 命令行，根本不是 spec 键；而 `SKILL.md`
> 当天正在增补小节，同一句话实测先在 L135、几十分钟后漂到 L246 —— 行号在这个文件里
> 不可靠。要定位就 `grep -n font_east_asia SKILL.md`。
>
> 上表的两个脚本行号同理，只是**当天**快照：`grep -n 'def set_fonts' scripts/fix_fonts.py`、
> `grep -n 'def _set_style_fonts' scripts/docx_create.py`。`docx_create.py:211` 这条已经
> 失效过一次（spec 校验新增 ~400 行后真身挪到 L222，L211 只剩一段 docstring 尾巴），
> `fix_fonts.py:89` 复核仍准。

**两条铁律跨入口都成立**——换个入口，约束不换：

- **`w:rFonts` 只能有一个**：`fix_fonts` 侧靠 `dedupe_rpr()`；`docx_create` 侧钉在
  `test_existing_rfonts_is_extended_not_duplicated`（已有 `rFonts` 就地更新，不追加第二个）。
- **`w:rFonts` 必须是 `w:rPr` 的第一个子元素**：`fix_fonts` 侧就是本文铁律 2 的 `rpr.insert(0, rf)`；
  `docx_create` 侧交给 `rpr.insert_element_before(rf, "w:b", …)`，由 python-docx 的 `CT_RPr`
  序列表落位，钉在 `test_rfonts_stays_before_later_rpr_children`。

**唯一的实现差异，别互相抄**：`fix_fonts.py` 是 lxml 裸元素管线，只有 `insert(0, …)` 一条路；
`docx_create.py` 跑在 python-docx 的 `CT_RPr` 上，必须走 `insert_element_before()` 让库排序列——
在那里写 `insert(0, …)` 会插到 `rStyle` 前面，schema 非法。

**改任一侧就同步另一侧**：槽位约定（`ascii`/`hAnsi`/`cs` 恒 TNR，`eastAsia` 单独给）与顺序约定
是**一份语义、两个实现**；改一处要连带改另一处和各自的测试，否则两个入口会在同一份文档上
给出不同结果。

## 两条铁律

违反任意一条，Word 直接报「文件可能已损坏」。

**铁律 1：`w:rPr` 内同名子元素只能留一个，先 dedupe 只留最后一个。** `rPr` 子元素是
`xsd:sequence` 且多数 `maxOccurs=1`，重复即 schema 非法；实测 Word/WPS 打开含重复 `rFonts` 会
直接弹修复提示。

```python
def dedupe_rpr(rpr):
    by_tag = {}
    for child in list(rpr):
        by_tag.setdefault(child.tag, []).append(child)
    for tag, els in by_tag.items():
        for e in els[:-1]:          # 只留最后一个
            rpr.remove(e)
```

**铁律 2：`w:rFonts` 必须是 `w:rPr` 的第一个子元素。** 它是 CT_RPr 序列首位，删旧 `rFonts`
后必须 **`rpr.insert(0, rf)` 而非 `append`**（后者到尾部，Word 判非法）：

```python
def set_fonts(rpr, ea):
    for old in rpr.findall(q("rFonts")):
        rpr.remove(old)
    rf = etree.Element(q("rFonts"))
    for a in ("ascii", "hAnsi", "cs"):
        rf.set(q(a), "Times New Roman")
    rf.set(q("eastAsia"), ea)
    rpr.insert(0, rf)                       # 铁律 2：不是 append
```

它自身不会造成重复，但**调用前仍要 dedupe**——文档可能带多个 `rFonts`，`findall` 只删这一种。
顺序固定：`dedupe_rpr` → `set_fonts`。

## `styles[]` 命中已存在样式：就地 modify，只改给出的键

`add_styles()`（`scripts/docx_create.py:355`，取样式入口 `_get_or_add_style()` 在 `:340`）现在的语义是**命中即就地改**：

```python
style, created = _get_or_add_style(doc, s["name"])   # try doc.styles[name] / except KeyError: add_style
```

| 情况 | 行为 |
|---|---|
| 名字命中已存在样式（含内置 `Normal` / `Heading 1`） | 拿到**同一个** `w:style` 元素就地改，`after.element is before.element` |
| 名字不存在 | `add_style` 新建，行为与从前一致 |
| 同名条目二次 apply | `styles.xml` 里该 `w:style` 恒 1 个，不累积 |

**「只改给出的键」是硬要求。** `size_pt` 单独给时，`font` / `bold` / `italic` / `base` / `outlineLvl` / 主题字体绑定**逐字保留**——否则 `{"name": "Heading 1", "size_pt": 16}` 会把上文中和出来的黑色冲回主题蓝，那正是 P12 之前反复踩的坑。四个方向都有测试钉住：不给 `bold` / 不给 `base` / 不给 `outline_level` / 不给 `italic`。

**想反悔颜色就显式写 `color`。** 中和层（下一节）跑在 `add_styles` **之前**，所以 spec 的 `color` 仍然赢；`ColorFormat.rgb` 的 setter 内部先 `_remove_color()`，`themeColor` 一并消失——若直接改 `@w:val`，主题属性会继续赢。

**一条 exit 2 路径因此消失**：早先「同名样式已存在」会抛 `ValueError`（在 `_BUILD_ERRORS` 里），现在变成功。这是预期的行为变更，不是缺陷；`ensure_toc_styles` 是 `try/except KeyError` 自兜，从不依赖那次崩溃。

## 内置标题默认是黑的（`docx_create.py` 的中和层）

生成路径上还有一层与字体同源、但更靠前的处理：**`Heading 1`–`Heading 9` / `Title` / `Subtitle`
在套用任何 `styles[]` 之前先被中和成黑色**（`_neutralise_heading_styles()`，`scripts/docx_create.py:709`，
在 `main()` 的 `:1919` 调用，位置在 `add_styles` **之前**）。

它做的事，逐项：

| 目标 | 动作 |
|---|---|
| `w:color` | 整元素替换成 `w:val="000000"` —— `themeColor` / `themeShade` / `themeTint` 一并消失，不是「再加一个 `w:val`」 |
| 字符变体 | `Heading1Char` … `Heading9Char`、`TitleChar`、`SubtitleChar` 同样处理 |
| `Title` 的段落下边框 | `w:pBdr` 上写 `w:color="000000"`；`w:sz`=8、`w:space`=4 **不变** |
| `Subtitle` | 删 `w:i` 与 `w:iCs`（去斜体）；`Heading 9` 的斜体**保留** |
| 已显式着色的样式 | 不动，只在报告里记一笔（`explicit`） |
| 无 `w:rPr` 的样式 | 不凭空造颜色，记 `no_colour` |
| 中文模板名 | `标题 N` / `标题` / `副标题` 一并匹配；styleId 也匹配，所以本地化拼写不会让中和静默失效 |

**为什么必须先中和 `w:color` 而不是只写 `w:val`**：`CT_Color` 里 `themeColor` 优先级高于 `w:val`，
留着 `themeColor="accent1"` 再写 `w:val="000000"`，Word 渲染出来仍是主题蓝——XML 层看着改了，
屏幕上没变，是最典型的假绿灯。

**这不违反「`styles[]` 只改给出的键」**：中和是**默认值**，跑在前面；spec 里给了 `color` 的条目
照样覆盖（`add_styles` 的 `ColorFormat.rgb` setter 内部会 `_remove_color()`，主题属性一起清掉）。
所以「想反悔颜色就显式写 `color`」这条通道是真的通的，不是纸面承诺。

**同源的另一个坑：命名一个字面字体时必须摘掉主题字体绑定。** `_set_style_fonts()`（`scripts/docx_create.py:241`）写显式字体前会删 `w:asciiTheme` / `w:hAnsiTheme` / `w:eastAsiaTheme` / `w:cstheme`（`:270` 的循环）——因为 `CT_Fonts` 里主题属性优先于显式同名属性，只写 `w:ascii="黑体"` 会得到「XML 是黑体、Word 渲染主题字体」。这属于 honour `font` 的实现细节，不是改了 spec 没提的键：**没给 `font` 时一行都不动**。

## 中文排版默认值

`docx_utils` 三组常量，`ascii`/`hAnsi`/`cs` 恒为 TNR，只有 `eastAsia` 不同：`FONT_TNR` →
`微软雅黑`、`FONT_HEITI` → `黑体`、`FONT_SONG` → `宋体`。口径：**西文/数字一律 TNR，中文交给
eastAsia**。`FONT_TNR` 的 `eastAsia` 是**微软雅黑**——它指"西文用 TNR"这套组合；`fix_fonts.py`
只下发宋体（正文）与微软雅黑（标题）。

## fix_fonts.py 的三层判定

`python scripts/fix_fonts.py <docx> [<docx2> ...] [--heading-threshold 14]`，就地改写（临时文件 +
`shutil.move`，中断不损原文件）。**在仓库根执行**——`scripts/` 前缀不能省，裸写
`python fix_fonts.py` 会因为 `fix_fonts.py` 不在 cwd 而 `can't open file` 退出 1。

**第 1 层 样式层** —— 遍历 `styles.xml` 的 `w:style`：`Normal` → 宋体；`heading*` 或 `Title` →
雅黑；其余不碰。`docDefaults` 存在则 `dedupe_rpr` + `set_fonts(…, "宋体")`。

**第 2 层 run 层** —— 遍历每个 `w:p`，按 `pPr/pStyle` 判是否 heading 后对 `p.findall(w:r)`
分派：公式 run（`m:oMath` 内）只改 `ascii`/`hAnsi`/`cs` = TNR，**`eastAsia` 不动**（仅缺失才补
TNR）——公式中文与字形由 OMML 自管；普通 run 取 `微软雅黑`（heading）或 `宋体`。公式归属靠
**id 集合**判定，必须在遍历 run **之前**建好（嵌套 `doc.iter` 会让同一元素进多次集合，`id()`
去重正合适）：

```python
math_ids = {id(el) for om in doc.iter(f"{{{M}}}oMath") for el in om.iter()}
```

**第 3 层 手动大标题（v3）** —— 给"没套标题样式、只是手动放大字号"的段兜底：`rPr/w:sz` 的
`w:val`（**半磅**）≥ `threshold_half_pt`，**且段落无 `w:pStyle`** → 雅黑。`--heading-threshold`
单位 **pt**（默认 `14`），内部 ×2 换半磅。「无 `pStyle`」是判据核心——套了样式的段归第 1 层管：

```python
has_style = ppr is not None and ppr.find(q("pStyle")) is not None
is_manual = int(sz.get(q("val"))) >= threshold_half_pt and not has_style
```

`w:sz` 类型不对时静默跳过。收尾打一行 `正文 N  标题 M  手动大标题 K  公式 J`。

## 解包自检

```python
def check(path):                    # 期望 (0, 0, N)
    dup = not_first = total = 0
    with zipfile.ZipFile(path) as z:
        for part in ("word/document.xml", "word/styles.xml"):
            for rpr in etree.fromstring(z.read(part)).iter(q("rPr")):
                tags = [k.tag for k in rpr]
                n = tags.count(q("rFonts"))
                dup += max(0, n - 1)
                total += n
                not_first += n - (1 if tags and tags[0] == q("rFonts") else 0)
    return dup, not_first, total
```

判据：`duplicate_rFonts == 0` 且 `not_first_child == 0`，**不要只信 stdout**。**幂等**：连跑两次
后两条仍成立且 `rFonts` 总数不增。

## 几个配套的坑

- **zip 就地改写**必须"全部读入内存 → 写临时文件 → `shutil.move`"；边读边写会截断自己。
- **`standalone=True`** 是 Word 原文件形态，缺了不致命但能让 diff 最小。
- **run 里没 `rPr`** 时新建并插首位：`rpr = etree.Element(q("rPr")); r.insert(0, rpr)`。
- `mkrun()` 用 `rpr.append(rf)`（从零新建，追加即首位）；**修既有文档必须用 `insert(0, …)`**。

## 和公式链的交界

`math_omml.apply_math_font()` 也往 `<m:r>` 里塞 `w:rPr/w:rFonts`，但守的是 **CT_R** 顺序
（`m:rPr? → w:rPr? → m:t`）：`w:r` 内要求 `rFonts` 是 `w:rPr` 首个子元素；`m:r` 内要求 `w:rPr`
排在 `m:t` 之前。别在 `fix_fonts` 里对公式 run 做 `insert(0, rFonts)` 式全量重写，那会破坏
`m:r` 的顺序。