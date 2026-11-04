# 驗證 UnicodeMath：交給 Word 建構（BuildUp）成方程式，再轉 PDF 檢查
# 用法：node -e "..." 產生 um.json（latex/um 陣列）後執行本腳本 um.json 輸出.pdf
param([string]$Json, [string]$Pdf)
$items = Get-Content $Json -Raw -Encoding UTF8 | ConvertFrom-Json
$w = New-Object -ComObject Word.Application
$w.Visible = $false
$d = $w.Documents.Add()
$i = 0
foreach ($it in $items) {
  $i++
  $r = $d.Content
  $r.InsertParagraphAfter()
  $p = $d.Paragraphs($d.Paragraphs.Count).Range
  $p.Text = "$i. "
  $r = $d.Content
  $r.InsertParagraphAfter()
  $p = $d.Paragraphs($d.Paragraphs.Count).Range
  $p.Text = $it.um
  $m = $d.OMaths.Add($p)
  $m.OMaths($m.OMaths.Count).BuildUp() | Out-Null
}
"OMaths: " + $d.OMaths.Count
$d.ExportAsFixedFormat($Pdf, 17)
$d.Close($false)
$w.Quit()
