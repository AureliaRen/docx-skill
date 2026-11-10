# pdf_export.ps1 —— docx → PDF 导出（Word COM）
# MIT License. Part of the docx skill scripts.
#
# 解决什么问题
# ------------
# 把带 OMML 公式的 docx 导出成 PDF，且**文本层必须保留**（公式附近文字能搜到）。
#
# 为什么不用 WPS（硬约束 #43）
#   WPS 导出带 OMML 公式的 docx → PDF **文本层残缺**：公式变成图或整段被丢，
#   PDF 里搜不到公式附近文字。生产验证过一次（10 份 40 页 A4、463 个 OMML 的
#   交付物全部走 Word COM 导出）。Word COM 的 `ExportAsFixedFormat(17)`
#   （17 = wdExportFormatPDF）保留完整文本层。
#
# 用法
# ----
#   .\pdf_export.ps1 -DocxPath <path.docx>              # 输出同目录同名 .pdf
#   .\pdf_export.ps1 -DocxPath <path.docx> -PdfPath <out.pdf>
#
# 坑
# --
# - `Documents.Open` 第三个参数 $true = ReadOnly：只读打开不占用原文件，
#   也不给原 docx 写"最近打开"之类的脏数据。
# - COM 对象必须 try/finally 兜底：任何一步抛异常都要 `Quit()`，否则留僵尸
#   WINWORD 进程占文件句柄（任务管理器里杀不干净就锁目录）。
# - `$doc.Close($false)` 的 $false = 不保存；导出动作本身不修改文档。
# - Word 路径必须是绝对路径（COM 对相对路径的解析依赖进程工作目录，不可靠）。

param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath,
    [string]$PdfPath
)

$ErrorActionPreference = 'Stop'

$docxFull = [System.IO.Path]::GetFullPath($DocxPath)
if (-not [System.IO.File]::Exists($docxFull)) {
    Write-Error "docx not found: $docxFull"
    exit 1
}
if (-not $PdfPath) {
    $PdfPath = [System.IO.Path]::ChangeExtension($docxFull, '.pdf')
}
$pdfFull = [System.IO.Path]::GetFullPath($PdfPath)

Write-Host "docx: $docxFull"
Write-Host "pdf : $pdfFull"

$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Open($docxFull, $false, $true)   # 第三个 $true = ReadOnly
    $doc.ExportAsFixedFormat($pdfFull, 17)                   # 17 = wdExportFormatPDF
    Write-Host "exported OK"
}
finally {
    if ($doc)  { try { $doc.Close($false) } catch { Write-Warning "Close failed: $_" } }
    if ($word) { try { $word.Quit() }       catch { Write-Warning "Quit failed: $_" } }
    # 释放 COM 引用，帮 GC 收尾（不释放偶尔会留半死的 WINWORD）
    if ($doc)  { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
}

if (-not [System.IO.File]::Exists($pdfFull)) {
    Write-Error "PDF was not created: $pdfFull"
    exit 1
}
Write-Host "pdf_export done"
