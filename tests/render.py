"""用 Word 實際排版 make-docx.js 的輸出，轉 PDF → PNG 檢查。

用法：node tests/make-docx.js out.json && python tests/render.py out.json 輸出資料夾
需要 Windows + Word、Python 套件 pypdfium2。
"""
import json
import subprocess
import sys
import zipfile
from pathlib import Path

import pypdfium2 as pdfium

data = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
outdir = Path(sys.argv[2]).resolve()
outdir.mkdir(parents=True, exist_ok=True)

NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def text(s):
    return f'<w:r><w:t xml:space="preserve">{s}</w:t></w:r>'


paras = []
for i, d in enumerate(data, 1):
    paras.append(f'<w:p>{text(f"{i}. 行內：已知 ")}{d["inline"]}{text("，求值。")}</w:p>')
    paras.append(f'<w:p>{text("   Word 預設（會縮）：")}{d["shrink"]}{text("。")}</w:p>')
    paras.append(f'<w:p>{d["display"]}</w:p>')

doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
       + "".join(paras) +
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1000" w:bottom="1000" w:left="1000"/></w:sectPr>'
       '</w:body></w:document>')
files = {
    "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>',
    "_rels/.rels": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    "word/_rels/document.xml.rels": '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
    "word/styles.xml": '<?xml version="1.0" encoding="UTF-8"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="新細明體"/>'
    '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>',
    "word/document.xml": doc,
}
docx = outdir / "samples.docx"
with zipfile.ZipFile(docx, "w", zipfile.ZIP_DEFLATED) as z:
    for n, s in files.items():
        z.writestr(n, s)

pdf = outdir / "samples.pdf"
ps = (f"$w=New-Object -ComObject Word.Application;$d=$w.Documents.Open('{docx}',$false,$true);"
      f"'oMath:'+$d.OMaths.Count;$d.ExportAsFixedFormat('{pdf}',17);$d.Close($false);$w.Quit()")
print(subprocess.run(["powershell", "-NoProfile", "-Command", ps], capture_output=True, text=True).stdout.strip())
for i, page in enumerate(pdfium.PdfDocument(str(pdf))):
    page.render(scale=1.6).to_pil().save(outdir / f"page{i + 1}.png")
    print("page", i + 1)
