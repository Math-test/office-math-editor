// 把 manifest.xml 裡的網址換成你架設的位置（GitHub Pages）
// 用法：node tools/set-base-url.js https://你的帳號.github.io/office-math-editor
var fs = require('fs');
var path = require('path');

var base = (process.argv[2] || '').replace(/\/+$/, '');
if (!/^https:\/\/[^/]+/.test(base)) {
  console.error('請給 https 網址，例如：node tools/set-base-url.js https://my-name.github.io/office-math-editor');
  process.exit(1);
}
var file = path.join(__dirname, '..', 'manifest.xml');
var xml = fs.readFileSync(file, 'utf8');
// 找出目前的 base（taskpane 網址去掉 /src/taskpane.html）
var m = /<SourceLocation DefaultValue="([^"]+)\/src\/taskpane\.html"/.exec(xml);
if (!m) { console.error('manifest.xml 裡找不到 taskpane 網址'); process.exit(1); }
var oldBase = m[1];
var oldOrigin = /^https:\/\/[^/]+/.exec(oldBase)[0];
var newOrigin = /^https:\/\/[^/]+/.exec(base)[0];
xml = xml.split(oldBase).join(base).split('<AppDomain>' + oldOrigin + '</AppDomain>').join('<AppDomain>' + newOrigin + '</AppDomain>');
fs.writeFileSync(file, xml, 'utf8');
console.log('已更新 manifest.xml：' + oldBase + ' → ' + base);
