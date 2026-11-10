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
| 绑定函数 | `set_fonts(rpr, ea)`（`scripts/fix_fonts.py:89`）——删旧 `w:rFonts` 后 `rpr.insert(0, rf)` | `_set_style_fonts(style, latin, east_asia)`（`scripts/docx_create.py:211`）——`rpr.insert_element_before(rf, "w:b", …)` |
| 怎么表达 | CLI 参数 `--heading-threshold` + 三层判定 | spec 键 `font` / `font_east_asia`（`SKILL.md` L112） |
| 测试锚点 | 本文「解包自检」的 `(dup, not_first, total)` 判据 | `tests/test_style_fonts.py`（5 条） |

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

## 中文排版默认值

`docx_utils` 三组常量，`ascii`/`hAnsi`/`cs` 恒为 TNR，只有 `eastAsia` 不同：`FONT_TNR` →
`微软雅黑`、`FONT_HEITI` → `黑体`、`FONT_SONG` → `宋体`。口径：**西文/数字一律 TNR，中文交给
eastAsia**。`FONT_TNR` 的 `eastAsia` 是**微软雅黑**——它指"西文用 TNR"这套组合；`fix_fonts.py`
只下发宋体（正文）与微软雅黑（标题）。

## fix_fonts.py 的三层判定

`python fix_fonts.py <docx> [<docx2> ...] [--heading-threshold 14]`，就地改写（临时文件 +
`shutil.move`，中断不损原文件）。

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