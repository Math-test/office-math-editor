# 用 Word COM 模擬增益集在 Word 裡的流程
# （COM 的 WordOpenXML / InsertXML 與 Office.js 的 getOoxml / insertOoxml 是同一套機制）
#  1) 在句子中間插入方程式：佔位文字 → 整段 OOXML 換掉 → 寫回，不能多出段落
#  2) Word 原生行內方程式（老師用 Alt+= 打的）→ 放大分數 → 寫回
# 用法：powershell -File tests\word-flow.ps1 輸出資料夾
param([string]$Out)
$ErrorActionPreference = 'Stop'
$node = 'C:\Program Files\nodejs\node.exe'
$root = (Split-Path $PSScriptRoot) -replace '\\', '/'
$utf8 = New-Object Text.UTF8Encoding $false
New-Item -ItemType Directory -Force $Out | Out-Null

function RunNode($script, $arg1) {
  $js = "var fs=require('fs'),D=require('$root/node_modules/@xmldom/xmldom');" +
        "function P(s){return new D.DOMParser().parseFromString(s,'text/xml')}" +
        "function S(d){return new D.XMLSerializer().serializeToString(d)}" + $script
  & $node -e $js $arg1
}

$w = New-Object -ComObject Word.Application
$w.Visible = $false
$d = $w.Documents.Add()
$d.Content.Font.Size = 12
$d.Content.Text = "Word 原生（Alt+= 打的）：已知 x=(a+b)/c，又 y=1/2+3/4，求值。"
foreach ($lin in @('x=(a+b)/c', 'y=1/2+3/4')) {
  $f = $d.Paragraphs(1).Range.Duplicate
  $f.Find.Execute($lin) | Out-Null
  $m = $d.OMaths.Add($f)
  $m.OMaths(1).BuildUp() | Out-Null
}
$d.Content.InsertParagraphAfter()
$d.Paragraphs(2).Range.Text = "上面這段按「放大分數」後：已知 x=(a+b)/c，又 y=1/2+3/4，求值。"
foreach ($lin in @('x=(a+b)/c', 'y=1/2+3/4')) {
  $f = $d.Paragraphs(2).Range.Duplicate
  $f.Find.Execute($lin) | Out-Null
  $m = $d.OMaths.Add($f)
  $m.OMaths(1).BuildUp() | Out-Null
}
$d.Content.InsertParagraphAfter()
$d.Paragraphs(3).Range.Text = "用增益集插入：已知【EQMARK7731】，求值，句子沒有斷行。"
$before = $d.Paragraphs.Count

# 1) 插入：段落 3 的佔位文字換成方程式
$p3 = $d.Paragraphs(3).Range
[IO.File]::WriteAllText("$Out\p3.xml", $p3.WordOpenXML, $utf8)
RunNode ("var I=require('$root/src/insert-inline.js');var doc=P(fs.readFileSync(process.argv[1],'utf8'));" +
  "var ok=I.replaceMarker(doc,I.MARKER,'x=\\frac{-b\\pm\\sqrt{b^2-4ac}}{2a}',{fontSize:12,parseXml:P});" +
  "fs.writeFileSync(process.argv[1]+'.out',S(doc),'utf8');console.log('佔位替換：'+ok)") "$Out\p3.xml"
$p3.InsertXML([IO.File]::ReadAllText("$Out\p3.xml.out", $utf8))

# 2) 放大段落 2 既有方程式的分數
$p2 = $d.Paragraphs(2).Range
[IO.File]::WriteAllText("$Out\p2.xml", $p2.WordOpenXML, $utf8)
RunNode ("var F=require('$root/src/fixfrac.js');var doc=P(fs.readFileSync(process.argv[1],'utf8'));" +
  "var n=F.apply(doc,{fontSize:12});fs.writeFileSync(process.argv[1]+'.out',S(doc),'utf8');console.log('放大分數：'+n+' 個')") "$Out\p2.xml"
$p2.InsertXML([IO.File]::ReadAllText("$Out\p2.xml.out", $utf8))

"段落數：$before → " + $d.Paragraphs.Count + "（應相同）；方程式數：" + $d.OMaths.Count + "（應為 5）"
$d.SaveAs2("$Out\flow.docx", 16)
$d.ExportAsFixedFormat("$Out\flow.pdf", 17)
$d.Close($false)
$w.Quit()
