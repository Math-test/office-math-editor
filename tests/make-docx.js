// 把 samples 轉成 OMML，輸出 JSON 給 render.py 組成 docx 給 Word 排版檢查
var L = require('../src/latex2omml.js');
var samples = require('./samples.js');
var out = samples.map(function (s) {
  return {
    latex: s,
    inline: L.toOmml(s, { fontSize: 12 }),
    shrink: L.toOmml(s, { fontSize: 12, noShrink: false }),
    display: L.toOmml(s, { display: true })
  };
});
require('fs').writeFileSync(process.argv[2], JSON.stringify(out), 'utf8');
