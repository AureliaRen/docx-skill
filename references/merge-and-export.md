# 合并与导出技术参考

> ## ⚠️ PDF 导出的定位：附加能力，不是本 skill 的主路由
>
> 本 skill 的主路由（`SKILL.md` 的 description）明确写着 **"Do NOT use for PDF"** —— 这条护栏保留，
> 因为把 Word 文档当 PDF 生成器用是错的（要 PDF 该用专门的 PDF 工具链）。
>
> `scripts/pdf_export.ps1` 是**在特定前提下**才成立的附加能力：
>
> - **前提一：文档里有 OMML 公式（原生 Office Math）**。这才是不用现成 PDF 工具链的理由 ——
>   WPS 导出带 OMML 的 docx 会**丢文本层**（公式变图或整段被丢，PDF 里搜不到公式附近文字）；
>   Word COM 的 `ExportAsFixedFormat(17)` 文本层完整。**生产实证**：40 页 A4、463 个 OMML 公式、
>   10 份 Word 真机导出 PDF 的交付物全部走这条路。
> - **前提二：Windows + 本机装了 Microsoft Word**。COM 依赖 Word，Linux/macOS 与无 Word 的机器
>   一律不可用；CI 的 ubuntu 腿必须跳过它。
>
> 若你要的只是「把 docx 转 PDF」而没有公式、或不在 Windows 上 —— **别用这个脚本**，
> 走通用的 PDF 工具链更合适。这里的方法学只解决「带公式的 docx 转 PDF 且文本层不能丢」这一个窄问题。

## merge_docs.py 七步

python-docx 做不了跨包合并（styles / numbering / media 关系各自独立），所以直接在 zip/XML 层
操作：

```
python scripts/merge_docs.py <base.docx> <src.docx> <out.docx>
```

base 是基底（页面设置与样式以它为准），src 整段追加。任一输入不存在 → exit 1
（`error: … not found`）。在仓库根执行，`scripts/` 前缀不能省。

**1 样式** —— 按 `w:styleId` 并集去重，base 优先：已有的不动，src 独有的整段 `deepcopy` 补进来。

**2 编号（裁定 3）** —— 按 `w:numId` 并集去重。这是原实现"docstring 承诺却零实现"的地方：
缺失时 src 的 `w:numPr` 指向不存在的 numId，Word 要么编号错乱要么静默丢编号。

- base 无 `numbering.xml` 而 src 有 → 整个文件搬，并补 `[Content_Types].xml` 的 Override
  （`PartName="/word/numbering.xml"` + `…wordprocessingml.numbering+xml`）与
  `word/_rels/document.xml.rels` 的关系条目（`Id="rIdNumImport"`）。
- 双方都有 → 补 src 独有的 `w:num`，并按 `abstractNumId` 把引用的 `w:abstractNum` 一起搬。
  **只搬 num 不搬 abstractNum，编号定义照样是断的。** 引用链只有一层。

**3 media 重映射** —— src 图片复制进 base，文件名撞了加 `_1` 后缀；序列化的 r: 属性（不只
`r:embed`）撞号的新分配 rId。去重必须查 base **全量** media——两份文档的 `rId1` 各指向不同图。

**4 搬 body / 5 分页符** —— 块级元素逐个 `deepcopy` 进 base body，途中改 rId；`sectPr` 丢弃。
随后在 src 内容前插一个 `<w:br w:type="page"/>` 段落。

**6 sectPr 统一用 base（裁定 2）** —— src 自带的 sectPr 丢弃，只留 base 的页面设置。原实现注释
写"页边距统一"、行为却采用 src 的——页面设置不同的两份合并后**版心静默变成 src 的**且无报错。
实现与 docstring 不一致时以"跑一遍"为准。

**7 Content_Types + zip** —— 新 media 扩展名补 `Default`（png/jpg/jpeg/gif/emf/wmf，未命中落
`application/octet-stream`）；zip 按 base 原条目顺序重写。

## 合并后自检

```python
import zipfile
from xml.etree import ElementTree as ET
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

with zipfile.ZipFile(base) as z:
    b = ET.fromstring(z.read("word/document.xml"))
with zipfile.ZipFile(out) as z:
    o = ET.fromstring(z.read("word/document.xml"))

bp = len(list(b.iter(W + "p")))
assert len(list(o.iter(W + "p"))) == bp * 2 + 1        # +1 = 分页符段
sects = list(o.iter(W + "sectPr"))
assert len(sects) == 1                                 # 裁定 2
bs = b.find(f"{W}body/{W}sectPr")
for tag in ("pgSz", "pgMar"):                          # 逐属性等于 base
    assert dict(bs.find(W + tag).attrib) == dict(sects[0].find(W + tag).attrib)
```

**media 守恒**：1 张图的 base 与自身合并后应为 2。产物再跑 `docx_validate.py`（`ok: true`）。

## PDF 导出：必须走 Word COM

```
pwsh scripts/pdf_export.ps1 -DocxPath <path.docx>                  # 输出同目录同名 .pdf
pwsh scripts/pdf_export.ps1 -DocxPath <path.docx> -PdfPath <out.pdf>
```

在仓库根执行，路径带 `scripts/`；裸写 `.\pdf_export.ps1` 会因为仓储根没有这个文件而
`术语 '.\pdf_export.ps1' 不会被识别为 cmdlet`。跑通的输出是四行：`docx: …` / `pdf : …` /
`exported OK` / `pdf_export done`，退出码 0。

**为什么不用 WPS（硬约束 #43）**：WPS 导出带 OMML 的 docx → PDF **文本层残缺**，公式变图或整段
被丢，搜不到公式附近文字。生产验证过一次：10 份 40 页 A4、463 个 OMML 的交付物曾在 WPS 上翻车。
Word COM 的 `ExportAsFixedFormat($pdfFull, 17)`（`17 = wdExportFormatPDF`）保留完整文本层。

```powershell
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open($docxFull, $false, $true)   # 第三个 $true = ReadOnly
$doc.ExportAsFixedFormat($pdfFull, 17)
```

四条必须照做：**第三个参数 `$true` = ReadOnly**（不占原文件、不写脏数据）；**COM 必须
`try/finally` 兜底**（任一步抛异常都要 `Quit()`，否则留僵尸 WINWORD 占句柄）；**`Close($false)`**
的 `$false` = 不保存；**路径必须绝对**。`finally` 里 `Close`/`Quit` 后还要
`ReleaseComObject($doc)`（`$word` 同理），不释放偶尔留半死 WINWORD。

## PIL 多页 PDF 的正确写法

**不能每页 `save(append=True)` 反复开合**——每次 append 都重写整个文件尾，页数一多就慢，且
2000 页量级直接报 `trailer loop found`（PDF trailer 指向自己）。一次性交出去：

```python
pages_img[0].save("<out.pdf>", "PDF", resolution=300.0,
                  save_all=True, append_images=pages_img[1:])
```

**只有第 0 页负责 `save`**，其余页经 `append_images` 一次交给 PIL，文件只写一次；`resolution`
决定物理尺寸；每页须是**同尺寸画布**（先按最大页尺寸建白底画布，再按
`min(max_w / img.width, max_h / img.height, 1.0)` 缩放贴图）；图片要 `.convert("RGB")`，
带 alpha 的 PNG 直接进 PDF 会出黑底。
