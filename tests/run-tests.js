// 測試：1) 程式碼是 ES5（Office 2016/2019 的 IE11 核心能跑）
//       2) LaTeX → OMML → LaTeX → OMML 來回一致
//       3) 讀得懂 Word 自己產生的方程式
//       4) 分數不縮小：行內最外層分數有放大字級，巢狀/上標裡的不重複放大
var fs = require('fs');
var path = require('path');
var acorn = require('acorn');
var DOMParser = require('@xmldom/xmldom').DOMParser;
var L = require('../src/latex2omml.js');
var O = require('../src/omml2latex.js');
var samples = require('./samples.js');

var fails = 0, passes = 0;
function ok(cond, name, detail) {
  if (cond) { passes++; return; }
  fails++;
  console.log('✗ ' + name + (detail ? '\n    ' + detail : ''));
}
function parseXml(s) { return new DOMParser().parseFromString(s, 'text/xml'); }
function wrap(omml) { return '<w:document ' + L.NAMESPACES + '><w:body><w:p>' + omml + '</w:p></w:body></w:document>'; }

// 1) ES5
fs.readdirSync(path.join(__dirname, '..', 'src')).filter(function (f) { return /\.js$/.test(f); }).forEach(function (name) {
  var f = 'src/' + name, p = path.join(__dirname, '..', f);
  try { acorn.parse(fs.readFileSync(p, 'utf8'), { ecmaVersion: 5 }); ok(true); }
  catch (e) { ok(false, f + ' 不是 ES5（IE11 會壞）', e.message); }
});

// 2) 來回轉換
samples.forEach(function (s) {
  var o1 = L.toOmml(s, { noShrink: false });
  var back = O.fromOoxml(wrap(o1), parseXml);
  ok(back.length === 1, '讀回 1 個方程式：' + s, JSON.stringify(back));
  if (!back.length) return;
  var o2 = L.toOmml(back[0].latex, { noShrink: false });
  ok(o1 === o2, '來回一致：' + s, '讀回的 LaTeX：' + back[0].latex);
  ok(parseXml(wrap(L.toOmml(s))).getElementsByTagName('parsererror').length === 0, 'XML 合法：' + s);
});

// 3) Word 原生方程式
var zipPath = path.join(__dirname, 'fixtures', 'word-native.docx');
if (fs.existsSync(zipPath)) {
  var docXml = require('child_process').execFileSync('powershell', ['-NoProfile', '-Command',
    'Add-Type -A System.IO.Compression.FileSystem;$z=[IO.Compression.ZipFile]::OpenRead("' + zipPath + '");' +
    '$e=$z.GetEntry("word/document.xml");$s=New-Object IO.StreamReader($e.Open(),[Text.Encoding]::UTF8);' +
    '[Console]::OutputEncoding=[Text.Encoding]::UTF8;$s.ReadToEnd();$z.Dispose()'], { encoding: 'utf8', maxBuffer: 1 << 24 });
  var native = O.fromOoxml(docXml, parseXml).map(function (x) { return x.latex; });
  var expect = [
    'x=\\frac{a+b}{c}',
    '\\sqrt{b^{2}-4ac}',
    '\\sin ^{2}{\\theta }',  // 只檢查開頭片段
    '\\sum_{k=1}^{n}',
    '\\left(\\frac{1}{100}\\right)^{b}',
    '\\lim_{x\\to 0}'
  ];
  console.log('Word 原生方程式讀回：');
  native.forEach(function (l) { console.log('   ' + l); });
  ok(native.length === 6, '讀到 6 個 Word 原生方程式', native.length);
  expect.forEach(function (frag, i) {
    var got = (native[i] || '').replace(/\s/g, '');
    ok(got.indexOf(frag.replace(/\s/g, '').replace(/[{}]/g, '')) >= 0 || got.replace(/[{}]/g, '').indexOf(frag.replace(/\s|[{}]/g, '')) >= 0,
      'Word 原生 #' + (i + 1), native[i]);
    // 讀回的 LaTeX 要能再轉回合法 OMML
    ok(parseXml(wrap(L.toOmml(native[i] || ''))).getElementsByTagName('parsererror').length === 0, 'Word 原生 #' + (i + 1) + ' 可再轉回');
  });
}

// 4) 分數不縮小
var big = L.toOmml('\\frac{1}{2}', { fontSize: 12 });
ok(/<w:sz w:val="34"\/>/.test(big), '12pt 行內分數 → 分子分母 17pt（×1.4）', big);
var small = L.toOmml('\\frac{1}{2}', { fontSize: 12, noShrink: false });
ok(!/<w:sz /.test(small), '關閉時不改字級');
var disp = L.toOmml('\\frac{1}{2}', { display: true });
ok(!/<w:sz /.test(disp), '獨立方程式不需放大');
var sup = L.toOmml('e^{\\frac{1}{2}}', { fontSize: 12 });
ok(!/<w:sz /.test(sup), '上標裡的分數不放大');
var nested = L.toOmml('\\frac{\\frac{1}{2}}{3}', { fontSize: 10.5 });
ok((nested.match(/w:val="29"/g) || []).length >= 3, '10.5pt 繁分數：外層放大一次（29 半點），內層沿用', nested);

// 5) 解析細節
function latexOf(s) { return O.fromOoxml(wrap(L.toOmml(s, { noShrink: false })), parseXml)[0].latex; }
ok(latexOf('\\log_2 8').indexOf('\\log_{2} 8') === 0 || /\\log_\{2\}\s*8/.test(latexOf('\\log_2 8')), '\\log_2 8 下標只取 2', latexOf('\\log_2 8'));
ok(/dx=/.test(latexOf('\\int_0^1 x^2\\,dx=1').replace(/\s/g, '')), '積分吃到 dx 為止、不重複', latexOf('\\int_0^1 x^2\\,dx=1'));
ok(latexOf('x\\in\\mathbb{R}').indexOf('\\mathbb{R}') >= 0, '\\mathbb{R}');
ok(latexOf('x^23') === 'x^{2}3', 'x^23 只上標 2', latexOf('x^23'));

// 6) 段落中插入：拆開 run、保留格式、游標標記
var I = require('../src/insert-inline.js');
var W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
var para = parseXml('<w:document ' + W_NS + '><w:body><w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:tab/>' +
  '<w:t>已知' + I.MARKER + '，求值</w:t></w:r></w:p></w:body></w:document>');
ok(I.replaceMarker(para, I.MARKER, '\\frac{1}{2}', { parseXml: parseXml, cursorMarker: 'CUR' }), '找到佔位文字');
var p0 = para.getElementsByTagName('w:p')[0];
var kidsNames = [];
for (var c = p0.firstChild; c; c = c.nextSibling) kidsNames.push(c.nodeName);
ok(kidsNames.join(',') === 'w:r,m:oMath,w:r,w:r', '段落結構：前文 run、方程式、游標 run、後文 run', kidsNames.join(','));
var runs = p0.getElementsByTagName('w:r');
ok(runs[0].textContent === '已知', '前文', runs[0].textContent);
ok(runs[runs.length - 1].textContent === '，求值', '後文', runs[runs.length - 1].textContent);
ok(runs[runs.length - 1].getElementsByTagName('w:tab').length === 0, 'tab 不會被複製到後文');
ok(runs[runs.length - 1].getElementsByTagName('w:b').length === 1, '後文保留粗體');
ok(/w:val="39"/.test(new (require('@xmldom/xmldom').XMLSerializer)().serializeToString(p0)), '14pt 段落 → 分數 19.5pt（39 半點）');
var edge = parseXml('<w:document ' + W_NS + '><w:body><w:p><w:r><w:t>' + I.MARKER + '</w:t></w:r></w:p></w:body></w:document>');
I.replaceMarker(edge, I.MARKER, 'x', { parseXml: parseXml });
ok(edge.getElementsByTagName('w:r').length === 0 && edge.getElementsByTagName('m:oMath').length === 1, '整段只有佔位文字時只留方程式');

console.log('\n通過 ' + passes + '，失敗 ' + fails);
process.exit(fails ? 1 : 0);
