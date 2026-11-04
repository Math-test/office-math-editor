/*
 * 工作窗格主程式。純 ES5：Office 2016/2019 的增益集跑在 IE11 核心裡。
 */
(function () {
  'use strict';

  var VERSION = '2026.11.04-5'; // 顯示在窗格最下面，方便確認是不是讀到舊快取
  var HOST = null; // 'Word' | 'Excel' | null（直接用瀏覽器開）
  var CURSOR = 'EQCURSOR4419'; // 插入後把游標放到方程式後面用
  var HISTORY_KEY = 'ome.history';
  var OPTIONS_KEY = 'ome.options';

  // lim 的下面可以打字：下方加字的記號，本體是 \lim（見下方「弧、方框、上方加字」）
  var LIM_PROXY = '\\textcolor{#1f232d}{\\lim\\textcolor{#1f232e}{}}';

  // MathQuill 0.10 視覺編輯不支援這些（實測），遇到就改用 LaTeX 文字編輯
  var TEX_ONLY = /\\begin|\\overparen|\\underparen|\\boxed|\\hat|\\widehat|\\tilde|\\widetilde|\\dot|\\ddot|\\overset|\\underset|\\stackrel|\\iint|\\iiint|\\mathbb|\\overbrace|\\underbrace|\\genfrac|\\dfrac|\\tfrac|\\cfrac|\\operatorname\{(?!gcd|exp)/;

  // ---------- 符號面板 ----------
  // cmd：MathQuill 結構（游標會進到裡面）；w：直接寫入；seq：依序執行；tex：只能用 LaTeX 編輯
  var PALETTE = [
    [ // 常用
      { l: 'a/b', t: '分數（也可以直接打 /）', cmd: '\\frac' },
      { l: '√', t: '根號', cmd: '\\sqrt' },
      { l: 'ⁿ√', t: 'n 次方根', cmd: '\\nthroot' },
      { l: 'x²', t: '上標（也可以打 ^）', cmd: '^' },
      { l: 'x₁', t: '下標（也可以打 _）', cmd: '_' },
      { l: '( )', t: '括號', cmd: '(' },
      { l: '[ ]', t: '中括號', cmd: '[' },
      { l: '|x|', t: '絕對值', cmd: '|' },
      { l: 'AB⃗', t: '向量', cmd: '\\overrightarrow' },
      { l: 'A̅B̅', t: '線段（上橫線）', cmd: '\\overline' },
      { l: '⫽', t: '平行（斜線，課本寫法）', w: '⫽' },
      { l: '∠', t: '角', w: '\\angle ' },
      { l: '△', t: '三角形', w: '\\triangle ' },
      { l: '°', t: '度', w: '^{\\circ}' },
      { l: '±', t: '正負', w: '\\pm ' },
      { l: '×', t: '乘', w: '\\times ' },
      { l: '÷', t: '除', w: '\\div ' },
      { l: '≤', t: '小於等於', w: '\\le ' },
      { l: '≥', t: '大於等於', w: '\\ge ' },
      { l: '≠', t: '不等於', w: '\\ne ' },
      { l: 'Σ', t: '連加', cmd: '\\sum' },
      { l: '∫', t: '積分', cmd: '\\int' },
      { l: 'lim', t: '極限（下面可以打字，打完按 → 離開）', seq: [['w', LIM_PROXY], ['keys', 2]], wide: true },
      { l: 'logₐ', t: '對數', seq: [['w', '\\log'], ['cmd', '_']], wide: true },
      { l: 'Cⁿₖ', t: '組合數', seq: [['w', 'C'], ['cmd', '_']] },
      { l: '文字', t: '在方程式裡打中文或一般文字', cmd: '\\text', wide: true }
    ],
    [ // 符號
      '\\pm:±', '\\mp:∓', '\\times:×', '\\div:÷', '\\cdot:·', '\\le:≤', '\\ge:≥', '\\ne:≠', '\\approx:≈', '\\equiv:≡',
      '\\sim:∼', '\\cong:≅', '\\propto:∝', '\\infty:∞', '\\angle:∠', '\\triangle:△', '^{\\circ}:°', "':′",
      '\\perp:⊥', '⫽:⫽', '\\therefore:∴', '\\because:∵', '\\in:∈', '\\notin:∉', '\\subset:⊂',
      '\\subseteq:⊆', '\\cup:∪', '\\cap:∩', '\\varnothing:∅', '\\Rightarrow:⇒', '\\Leftrightarrow:⇔', '\\to:→',
      '\\cdots:⋯', '\\%:%', 'ℝ:ℝ', 'ℕ:ℕ', 'ℤ:ℤ', 'ℚ:ℚ'
    ],
    [ // 希臘
      '\\alpha:α', '\\beta:β', '\\gamma:γ', '\\delta:δ', '\\epsilon:ε', '\\theta:θ', '\\lambda:λ', '\\mu:μ',
      '\\pi:π', '\\rho:ρ', '\\sigma:σ', '\\tau:τ', '\\phi:φ', '\\omega:ω', '\\Delta:Δ', '\\Sigma:Σ', '\\Omega:Ω', '\\Phi:Φ'
    ],
    [ // 函數
      { l: 'sin', w: '\\sin ', wide: true }, { l: 'cos', w: '\\cos ', wide: true }, { l: 'tan', w: '\\tan ', wide: true },
      { l: 'cot', w: '\\cot ', wide: true }, { l: 'sec', w: '\\sec ', wide: true }, { l: 'csc', w: '\\csc ', wide: true },
      { l: 'sin²', w: '\\sin^2', wide: true }, { l: 'log', w: '\\log ', wide: true }, { l: 'ln', w: '\\ln ', wide: true },
      { l: 'logₐ', seq: [['w', '\\log'], ['cmd', '_']], wide: true },
      { l: 'lim', seq: [['w', LIM_PROXY], ['keys', 2]], wide: true }, { l: 'exp', w: '\\exp ', wide: true },
      { l: 'f(x)', w: 'f\\left(x\\right)', wide: true }, { l: "f′(x)", w: "f'\\left(x\\right)", wide: true }
    ],
    [ // 進階：grid 用格子編輯；tex 是 MathQuill 不支援、只能用 LaTeX 的
      { l: '分段函數', grid: { env: 'cases', rows: 2, cols: 2, prefix: 'f(x)=' },
        tex: 'f(x)=\\begin{cases}x^2 & x\\ge 0\\\\-x & x<0\\end{cases}', wide: true },
      { l: '聯立方程式', grid: { env: 'cases', rows: 2, cols: 1 }, tex: '\\begin{cases}x+y=3\\\\x-y=1\\end{cases}', wide: true },
      { l: '矩陣 [ ]', grid: { env: 'bmatrix', rows: 2, cols: 2 }, tex: '\\begin{bmatrix}a & b\\\\c & d\\end{bmatrix}', wide: true },
      { l: '矩陣 ( )', grid: { env: 'pmatrix', rows: 2, cols: 2 }, tex: '\\begin{pmatrix}a & b\\\\c & d\\end{pmatrix}', wide: true },
      { l: '行列式', grid: { env: 'vmatrix', rows: 2, cols: 2 }, tex: '\\begin{vmatrix}a & b\\\\c & d\\end{vmatrix}', wide: true },
      { l: 'Σ 總和', t: '連加（上下限可以直接打）', cmd: '\\sum', wide: true },
      { l: '∏ 連乘', cmd: '\\prod', wide: true },
      // #1f232? 是記號，見下方「弧、方框、上方加字」
      { l: '弧 ⌒AB', seq: [['w', '\\textcolor{#1f2329}{}'], ['keys', 1]], tex: '\\overparen{AB}', wide: true },
      { l: '方框', seq: [['w', '\\textcolor{#1f232a}{}'], ['keys', 1]], tex: '\\boxed{x}', wide: true },
      { l: '上方加字', t: '先打上面的小字，按 → 再打下面（例如等號上面寫「定義」）',
        seq: [['w', '\\textcolor{#1f232b}{\\textcolor{#1f232c}{}}'], ['keys', 2]],
        tex: '\\overset{\\text{定義}}{=}', wide: true },
      { l: '下方加字', t: '先打下面的小字，按 ← 回到上面打本體',
        seq: [['w', '\\textcolor{#1f232d}{\\textcolor{#1f232e}{}}'], ['keys', 2]],
        tex: '\\underset{\\text{說明}}{=}', wide: true },
      { l: '組合 (n k)', cmd: '\\binom', wide: true },
      { l: '集合 {…}', w: '\\left\\{\\right\\}', wide: true }
    ]
  ];

  // 可以用格子編輯的環境；l / r 是格子左右畫的括號
  var GRID_ENVS = {
    cases: { name: '大括號（聯立、分段）', l: '{', r: '' },
    bmatrix: { name: '矩陣', l: '[', r: ']' },
    pmatrix: { name: '矩陣', l: '(', r: ')' },
    Bmatrix: { name: '矩陣', l: '{', r: '}' },
    vmatrix: { name: '行列式', l: '|', r: '|' },
    Vmatrix: { name: '行列式', l: '‖', r: '‖' },
    matrix: { name: '矩陣（無括號）', l: '', r: '' },
    aligned: { name: '多行對齊', l: '', r: '' }
  };

  var mq = null, mf = null, texMode = false, syncing = false;
  var grid = null;   // { env, cells: [[MathField]], r, c }：目前的格子與最後停留的格子
  var active = null; // 符號面板寫入的編輯框（主編輯框或某一格）
  var $ = function (id) { return document.getElementById(id); };

  // ---------- 工具 ----------
  function setStatus(text, cls) {
    var s = $('status');
    s.textContent = text || '';
    s.className = cls || '';
  }
  function parseXml(s) { return new DOMParser().parseFromString(s, 'text/xml'); }
  function serialize(doc) { return new XMLSerializer().serializeToString(doc); }
  function store(key, val) { try { window.localStorage.setItem(key, JSON.stringify(val)); } catch (e) { /* 無痕或被封鎖 */ } }
  function load(key, dflt) {
    try { var v = window.localStorage.getItem(key); return v ? JSON.parse(v) : dflt; } catch (e) { return dflt; }
  }
  function errText(e) {
    if (!e) return '未知錯誤';
    var msg = e.message || String(e);
    if (e.debugInfo && e.debugInfo.message) msg += '（' + e.debugInfo.message + '）';
    return msg;
  }

  // ---------- 編輯器 ----------
  function currentLatex() {
    if (texMode || !mf) return $('latex').value;
    return proxyToTex(mf.latex() + (grid ? gridLatex() : ''));
  }

  function setTexMode(on) {
    texMode = on;
    $('texNotice').className = on ? '' : 'hidden';
    $('mathField').style.opacity = $('gridBox').style.opacity = on ? '0.35' : '';
    if (on) $('texBox').setAttribute('open', 'open');
  }

  // 放進視覺編輯器（必要時加格子）；放不進去回傳 false
  function loadVisual(latex) {
    latex = texToProxy(latex);
    var g = splitEnv(latex);
    if (g) {
      if (TEX_ONLY.test(g.prefix) || TEX_ONLY.test(g.rows.join(' '))) g = null;
      else {
        mf.latex(g.prefix);
        if ((g.prefix.replace(/\s/g, '') && !mf.latex()) || !makeGrid(g.env, g.rows)) g = null;
        else return true;
      }
      if (!g) { removeGrid(); mf.latex(''); return false; }
    }
    removeGrid();
    if (TEX_ONLY.test(latex)) { mf.latex(''); return false; }
    mf.latex(latex);
    // MathQuill 讀不懂時會變空白
    return !latex.replace(/\s/g, '') || !!mf.latex();
  }

  function setLatex(latex) {
    syncing = true;
    $('latex').value = latex;
    setTexMode(mf ? !loadVisual(latex) : true);
    syncing = false;
    decorate();
  }

  function onMathEdit() {
    decorate();
    if (syncing || texMode) return;
    $('latex').value = currentLatex();
  }

  function onLatexInput() {
    if (!mf) return;
    syncing = true;
    var ok = loadVisual($('latex').value);
    syncing = false;
    setTexMode(!ok);
    decorate();
  }

  // ---------- 弧、方框、上方加字 ----------
  // MathQuill 不支援 \overparen 等。借用 \textcolor{#1f232?}{…} 當記號（跟字的顏色幾乎一樣），
  // decorate() 依實際顏色加 class 讓 CSS 畫出外觀；進出編輯器時再和真正的 LaTeX 互換
  // 上方加字是 \textcolor{ovs}{\textcolor{ovt}{上}下}：打完上面按 → 就到下面
  // 下方加字是 \textcolor{und}{本體\textcolor{unb}{下面}}；本體是 \lim 時轉成 \lim_{…}
  var PTAG = { arc: '#1f2329', box: '#1f232a', ovs: '#1f232b', ovt: '#1f232c', und: '#1f232d', unb: '#1f232e' };
  var PNAME = { '9': 'arc', a: 'box', b: 'ovs', c: 'ovt', d: 'und', e: 'unb' };       // 顏色最後一碼 → 名稱
  var PBLUE = { 41: 'arc', 42: 'box', 43: 'ovs', 44: 'ovt', 45: 'und', 46: 'unb' };   // rgb 藍色值 → 名稱
  var PROXY = { arc: '\\overparen', box: '\\boxed' };

  function decorate() {
    var els = document.querySelectorAll('#editorBox .mq-textcolor');
    for (var i = 0; i < els.length; i++) {
      if (els[i].className.indexOf('mq-p-') >= 0) continue;
      var m = /rgb\(31,\s*35,\s*(\d+)\)/.exec(window.getComputedStyle(els[i]).color);
      if (m && PBLUE[m[1]]) els[i].className += ' mq-p-' + PBLUE[m[1]];
    }
  }

  // s.charAt(i) 是 '{'：回傳 { body, end }，end 在對應的 '}' 後面
  function readGroup(s, i) {
    var depth = 0;
    for (var j = i; j < s.length; j++) {
      var ch = s.charAt(j);
      if (ch === '\\') { j++; continue; }
      if (ch === '{') depth++;
      else if (ch === '}' && --depth === 0) return { body: s.slice(i + 1, j), end: j + 1 };
    }
    return null;
  }

  // 一個參數：{…}、\指令 或單一字元
  function readArg(s, i) {
    while (s.charAt(i) === ' ') i++;
    var ch = s.charAt(i);
    if (ch === '{') return readGroup(s, i);
    if (ch === '\\') {
      var m = /^\\([a-zA-Z]+|.)/.exec(s.slice(i));
      return m ? { body: m[0], end: i + m[0].length } : null;
    }
    return ch ? { body: ch, end: i + 1 } : null;
  }

  // 在 s 裡找某個記號的內容
  function proxyBody(s, name) {
    var tag = '\\textcolor{' + PTAG[name] + '}', i = s.indexOf(tag);
    var g = i < 0 ? null : readGroup(s, i + tag.length);
    return g ? g.body : '';
  }

  function proxyToTex(s) {
    var re = /\\textcolor\{#1f232([9a-e])\}/gi, out = '', last = 0, m;
    while ((m = re.exec(s))) {
      var g = readGroup(s, m.index + m[0].length), name = PNAME[m[1].toLowerCase()], inner, base;
      if (!g) break;
      out += s.slice(last, m.index);
      if (name === 'ovs') {
        // {\textcolor{ovt}{上}下}：ovt 以外的內容都是下面那一行
        inner = proxyBody(g.body, 'ovt');
        base = g.body.replace('\\textcolor{' + PTAG.ovt + '}{' + inner + '}', '');
        out += '\\overset{' + proxyToTex(inner) + '}{' + proxyToTex(base) + '}';
      } else if (name === 'und') {
        inner = proxyBody(g.body, 'unb');
        base = g.body.replace('\\textcolor{' + PTAG.unb + '}{' + inner + '}', '');
        if (/^\s*\\lim\s*$/.test(base)) out += '\\lim_{' + proxyToTex(inner) + '}';
        else out += '\\underset{' + proxyToTex(inner) + '}{' + proxyToTex(base) + '}';
      } else out += (PROXY[name] || '') + '{' + proxyToTex(g.body) + '}';
      last = re.lastIndex = g.end;
    }
    return out + s.slice(last);
  }

  function texToProxy(s) {
    s = s.replace(/\\parallel(?![a-zA-Z])\s*/g, '⫽'); // 編輯器也顯示斜的平行
    var re = /\\(overparen|boxed|overset|underset|lim)(?![a-zA-Z])/g, out = '', last = 0, m;
    while ((m = re.exec(s))) {
      var at = m.index + m[0].length, a, b, rep;
      if (m[1] === 'lim') {
        // 只有寫了下標的 lim 才換
        var k = at;
        while (s.charAt(k) === ' ') k++;
        if (s.charAt(k) !== '_') continue;
        a = readArg(s, k + 1);
        if (!a) break;
        rep = '\\textcolor{' + PTAG.und + '}{\\lim\\textcolor{' + PTAG.unb + '}{' + texToProxy(a.body) + '}}';
        out += s.slice(last, m.index) + rep;
        last = re.lastIndex = a.end;
        continue;
      }
      a = readArg(s, at);
      if (!a) break;
      if (m[1] === 'overset' || m[1] === 'underset') {
        b = readArg(s, a.end);
        if (!b) break;
        rep = m[1] === 'overset'
          ? '\\textcolor{' + PTAG.ovs + '}{\\textcolor{' + PTAG.ovt + '}{' + texToProxy(a.body) + '}' + texToProxy(b.body) + '}'
          : '\\textcolor{' + PTAG.und + '}{' + texToProxy(b.body) + '\\textcolor{' + PTAG.unb + '}{' + texToProxy(a.body) + '}}';
        a = b;
      } else rep = '\\textcolor{' + (m[1] === 'boxed' ? PTAG.box : PTAG.arc) + '}{' + texToProxy(a.body) + '}';
      out += s.slice(last, m.index) + rep;
      last = re.lastIndex = a.end;
    }
    return out + s.slice(last);
  }

  // ---------- 格子（矩陣、聯立方程式、分段函數）----------
  // MathQuill 不支援 \begin…\end，改成一格一個小編輯框，列、欄用按鈕增刪

  // 'f(x)=\begin{cases}a & b\\c & d\end{cases}' → { prefix, env, rows: [['a','b'],['c','d']] }
  function splitEnv(latex) {
    var m = /\\begin\{([A-Za-z]+)\}/.exec(latex);
    if (!m || !GRID_ENVS[m[1]]) return null;
    var env = m[1], start = m.index + m[0].length, endTag = '\\end{' + env + '}';
    var end = latex.lastIndexOf(endTag);
    if (end < start) return null;
    var body = latex.slice(start, end);
    // 巢狀環境、後面還有東西：交給 LaTeX 編輯
    if (/\\begin\{/.test(body) || latex.slice(end + endTag.length).replace(/\s/g, '')) return null;
    var rows = [[]], cell = '', depth = 0, i, j;
    for (i = 0; i < body.length; i++) {
      var ch = body.charAt(i);
      if (ch === '\\') {
        if (depth === 0 && body.charAt(i + 1) === '\\') { rows[rows.length - 1].push(cell); rows.push([]); cell = ''; }
        else cell += ch + body.charAt(i + 1);
        i++;
        continue;
      }
      if (ch === '{') depth++;
      else if (ch === '}') depth--;
      else if (ch === '&' && depth === 0) { rows[rows.length - 1].push(cell); cell = ''; continue; }
      cell += ch;
    }
    rows[rows.length - 1].push(cell);
    var last = rows[rows.length - 1];
    if (rows.length > 1 && last.length === 1 && !last[0].replace(/\s/g, '')) rows.pop(); // 結尾多一個 \\
    var cols = 1;
    for (i = 0; i < rows.length; i++) {
      for (j = 0; j < rows[i].length; j++) rows[i][j] = rows[i][j].replace(/^\s+|\s+$/g, '');
      cols = Math.max(cols, rows[i].length);
    }
    for (i = 0; i < rows.length; i++) while (rows[i].length < cols) rows[i].push('');
    return { prefix: latex.slice(0, m.index), env: env, rows: rows };
  }

  function gridLatex() {
    var rows = [];
    for (var i = 0; i < grid.cells.length; i++) {
      var r = [];
      for (var j = 0; j < grid.cells[i].length; j++) r.push(grid.cells[i][j].latex());
      rows.push(r.join(' & '));
    }
    return '\\begin{' + grid.env + '}' + rows.join('\\\\') + '\\end{' + grid.env + '}';
  }

  function snapshot() {
    var rows = [];
    for (var i = 0; i < grid.cells.length; i++) {
      rows.push([]);
      for (var j = 0; j < grid.cells[i].length; j++) rows[i].push(grid.cells[i][j].latex());
    }
    return rows;
  }

  function emptyRows(r, c) {
    var rows = [];
    for (var i = 0; i < r; i++) { rows.push([]); for (var j = 0; j < c; j++) rows[i].push(''); }
    return rows;
  }

  function bracketCell(tr, chr, n) {
    var td = tr.insertCell(-1);
    td.rowSpan = n;
    td.className = 'br';
    td.style.fontSize = (n * 34) + 'px';
    td.textContent = chr;
  }

  function fieldConfig(handlers) {
    return {
      spaceBehavesLikeTab: false,
      leftRightIntoCmdGoes: 'up',
      restrictMismatchedBrackets: true,
      sumStartsWithNEquals: false, // 高中多用 k，不要預先放 n=
      supSubsRequireOperand: true,
      charsThatBreakOutOfSupSub: '+-=<>',
      autoSubscriptNumerals: false,
      autoCommands: 'pi theta sqrt sum prod int alpha beta gamma delta lambda mu sigma omega phi infty pm times div le ge ne approx angle',
      autoOperatorNames: 'sin cos tan cot sec csc log ln lim exp max min',
      handlers: handlers
    };
  }

  // 每一格：Enter 加一列、方向鍵換格、在空白列開頭按 Backspace 刪掉這一列
  function cellHandlers(i, j) {
    return {
      edit: onMathEdit,
      enter: function () { addRow(i + 1); },
      moveOutOf: function (dir) {
        var cols = grid.cells[0].length;
        var k = i * cols + j + (dir === mq.L ? -1 : 1);
        if (k >= 0 && k < grid.cells.length * cols) focusCell(Math.floor(k / cols), k % cols);
        else if (dir === mq.L) mf.focus();
      },
      upOutOf: function () { if (i > 0) focusCell(i - 1, j); else mf.focus(); },
      downOutOf: function () { if (i < grid.cells.length - 1) focusCell(i + 1, j); },
      deleteOutOf: function (dir) {
        if (dir !== mq.L) return;
        var row = grid.cells[i];
        for (var k = 0; k < row.length; k++) if (row[k].latex()) { if (j > 0) focusCell(i, j - 1); return; }
        if (grid.cells.length > 1) { delRow(i); focusCell(Math.max(0, i - 1), grid.cells[0].length - 1); }
        else { removeGrid(); onMathEdit(); mf.focus(); }
      }
    };
  }

  // 依 rows（LaTeX 字串二維陣列）重畫格子；有格子 MathQuill 讀不懂就回傳 false
  function makeGrid(env, rows) {
    var table = $('gridTable'), info = GRID_ENVS[env], ok = true, was = syncing;
    syncing = true;
    table.innerHTML = '';
    grid = { env: env, cells: [], r: 0, c: 0 };
    for (var i = 0; i < rows.length; i++) {
      var tr = table.insertRow(-1);
      if (i === 0 && info.l) bracketCell(tr, info.l, rows.length);
      grid.cells.push([]);
      for (var j = 0; j < rows[i].length; j++) {
        var span = document.createElement('span');
        span.className = 'cell';
        tr.insertCell(-1).appendChild(span);
        var f = mq.MathField(span, fieldConfig(cellHandlers(i, j)));
        f.latex(rows[i][j]);
        if (rows[i][j].replace(/\s/g, '') && !f.latex()) ok = false;
        span.addEventListener('focusin', (function (r, c, fld) {
          return function () { grid.r = r; grid.c = c; active = fld; };
        })(i, j, f));
        grid.cells[i].push(f);
      }
      if (i === 0 && info.r) bracketCell(tr, info.r, rows.length);
    }
    $('gridTitle').textContent = info.name;
    $('gridAddCol').className = $('gridDelCol').className = (env === 'cases' ? 'hidden' : '');
    $('gridBox').className = '';
    syncing = was;
    return ok;
  }

  function removeGrid() {
    grid = null;
    active = mf;
    $('gridTable').innerHTML = '';
    $('gridBox').className = 'hidden';
  }

  function focusCell(r, c) {
    if (!grid) return;
    r = Math.max(0, Math.min(r, grid.cells.length - 1));
    c = Math.max(0, Math.min(c, grid.cells[r].length - 1));
    grid.cells[r][c].focus();
    grid.r = r; grid.c = c; active = grid.cells[r][c];
  }

  function regrid(rows, r, c) {
    if (!rows.length || !rows[0].length) { removeGrid(); onMathEdit(); mf.focus(); return; }
    makeGrid(grid.env, rows);
    onMathEdit();
    focusCell(r, c);
  }
  function addRow(at) {
    var rows = snapshot();
    rows.splice(at, 0, emptyRows(1, rows[0].length)[0]);
    regrid(rows, at, 0);
  }
  function delRow(at) {
    var rows = snapshot();
    rows.splice(at, 1);
    regrid(rows, at, grid.c);
  }
  function addCol(at) {
    var rows = snapshot();
    for (var i = 0; i < rows.length; i++) rows[i].splice(at, 0, '');
    regrid(rows, grid.r, at);
  }
  function delCol(at) {
    var rows = snapshot();
    for (var i = 0; i < rows.length; i++) rows[i].splice(at, 1);
    regrid(rows, grid.r, at);
  }

  function applyPaletteItem(it) {
    if (it.grid && mf && !texMode) {
      if (grid) {
        setStatus('一個式子只能放一個矩陣／大括號。要換掉的話，先按格子上方的「刪除整個」。', 'err');
        return;
      }
      if (it.grid.prefix && !mf.latex()) mf.latex(it.grid.prefix);
      makeGrid(it.grid.env, emptyRows(it.grid.rows, it.grid.cols));
      onMathEdit();
      focusCell(0, 0);
      return;
    }
    if (it.tex !== undefined && (texMode || !mf || !it.seq)) {
      var ta = $('latex');
      // 有內容就加在後面，沒有就直接放
      ta.value = ta.value ? ta.value + ' ' + it.tex : it.tex;
      setTexMode(true);
      $('texBox').setAttribute('open', 'open');
      ta.focus();
      return;
    }
    if (texMode || !mf) {
      var t = $('latex');
      t.value += it.w || it.cmd || '';
      return;
    }
    var target = (grid && active) || mf;
    var steps = it.seq || [[it.cmd ? 'cmd' : 'w', it.cmd || it.w]];
    for (var i = 0; i < steps.length; i++) {
      if (steps[i][0] === 'cmd') target.cmd(steps[i][1]);
      else if (steps[i][0] === 'keys') { for (var k = 0; k < steps[i][1]; k++) target.keystroke('Left'); }
      else target.write(steps[i][1]);
    }
    target.focus();
  }

  function buildPalette(tab) {
    var box = $('palette');
    box.innerHTML = '';
    var items = PALETTE[tab];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (typeof it === 'string') {
        var sep = it.lastIndexOf(':');
        it = { w: it.slice(0, sep) + (it.charAt(sep - 1).match(/[a-z]/i) ? ' ' : ''), l: it.slice(sep + 1) };
      }
      var b = document.createElement('button');
      b.type = 'button';
      b.textContent = it.l;
      b.title = it.t || it.l;
      if (it.wide) b.className = 'wide';
      b.onclick = (function (item) { return function () { applyPaletteItem(item); }; })(it);
      box.appendChild(b);
    }
    var tabs = document.querySelectorAll('.tab');
    for (i = 0; i < tabs.length; i++) tabs[i].className = 'tab' + (+tabs[i].getAttribute('data-tab') === tab ? ' on' : '');
  }

  // ---------- 最近使用 ----------
  function pushHistory(latex) {
    var h = load(HISTORY_KEY, []);
    var idx = h.indexOf(latex);
    if (idx >= 0) h.splice(idx, 1);
    h.unshift(latex);
    if (h.length > 12) h.length = 12;
    store(HISTORY_KEY, h);
    renderHistory();
  }
  function renderHistory() {
    var h = load(HISTORY_KEY, []), box = $('history');
    box.innerHTML = '';
    if (!h.length) {
      var e = document.createElement('div');
      e.className = 'empty';
      e.textContent = '插入過的式子會出現在這裡，點一下就能再用。';
      box.appendChild(e);
      return;
    }
    for (var i = 0; i < h.length; i++) {
      var b = document.createElement('button');
      b.type = 'button';
      b.textContent = h[i];
      b.title = h[i];
      b.onclick = (function (l) { return function () { setLatex(l); if (mf && !texMode) mf.focus(); }; })(h[i]);
      box.appendChild(b);
    }
  }

  function options() {
    return { noShrink: $('optNoShrink').checked, display: $('optDisplay').checked, italicNum: $('optItalicNum').checked };
  }

  // ---------- 複製 ----------
  function copyText(text) {
    var ta = $('clip');
    ta.value = text;
    ta.focus();
    ta.select();
    var ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    if (!ok && window.clipboardData) { // IE11
      try { ok = window.clipboardData.setData('Text', text); } catch (e2) { ok = false; }
    }
    return ok;
  }

  // ---------- Word ----------
  // 回傳 true 表示走了備用插入方式
  function wordInsert(latex) {
    var opt = options(), fallback = false;
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load('font/size');
      return ctx.sync().then(function () {
        var size = sel.font.size || 12;
        if (opt.display) {
          // 獨立一行：直接插入段落（Word 會自己斷行）
          sel.insertOoxml(Latex2Omml.toFlatOpc(latex, { display: true, italicNum: opt.italicNum }), 'Replace');
          return ctx.sync();
        }
        // 行內：佔位文字 → 整段 OOXML 換掉，避免 insertOoxml 把句子切成兩段
        var marker = sel.insertText(InsertInline.MARKER, 'Replace');
        var paras = marker.paragraphs;
        paras.load('items');
        return ctx.sync().then(function () {
          var para = paras.items[0];
          var ox = para.getOoxml();
          return ctx.sync().then(function () {
            var doc = parseXml(ox.value);
            var ok = InsertInline.replaceMarker(doc, InsertInline.MARKER, latex, {
              fontSize: size, noShrink: opt.noShrink, italicNum: opt.italicNum, parseXml: parseXml, cursorMarker: CURSOR
            });
            if (!ok) {
              // 段落裡讀不到佔位文字：退而求其次直接蓋掉佔位文字（Word 可能會多一個換行）
              marker.insertOoxml(Latex2Omml.toFlatOpc(latex, { italicNum: opt.italicNum }), 'Replace');
              fallback = true;
              return ctx.sync();
            }
            para.insertOoxml(serialize(doc), 'Replace');
            return ctx.sync();
          }).then(function () {
            // 游標移到方程式後面，方便接著打字
            var found = ctx.document.body.search(CURSOR, { matchCase: true });
            found.load('items');
            return ctx.sync().then(function () {
              if (found.items.length) {
                found.items[0].select();
                found.items[0].delete();
              }
              return ctx.sync();
            });
          });
        });
      });
    }).then(function () { return fallback; });
  }

  function wordLoad() {
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      var ox = sel.getOoxml();
      return ctx.sync().then(function () {
        var found = Omml2Latex.fromOoxml(ox.value, parseXml);
        if (!found.length) throw new Error('選取範圍裡沒有方程式。請拖曳選取整個方程式後再按。');
        setLatex(found[0].latex);
        $('optDisplay').checked = found[0].display;
        return found.length;
      });
    });
  }

  // 只處理選取範圍內「有行內方程式」的段落
  function wordFixFractions(undo) {
    return Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      sel.load('font/size');
      var all = sel.getOoxml();
      var paras = sel.paragraphs;
      paras.load('items');
      return ctx.sync().then(function () {
        var doc = parseXml(all.value);
        var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        var ps = doc.getElementsByTagNameNS(W, 'p');
        var targets = [];
        for (var i = 0; i < ps.length && i < paras.items.length; i++) {
          if (ps[i].getElementsByTagNameNS(Omml2Latex.M_NS, 'f').length) targets.push(paras.items[i]);
        }
        if (!targets.length) return 0;
        var oxs = targets.map(function (p) { return p.getOoxml(); });
        return ctx.sync().then(function () {
          var total = 0;
          for (var k = 0; k < targets.length; k++) {
            var d = parseXml(oxs[k].value);
            var n = FixFrac.apply(d, { fontSize: sel.font.size || 12, undo: undo });
            if (n) { targets[k].insertOoxml(serialize(d), 'Replace'); total += n; }
          }
          return ctx.sync().then(function () { return total; });
        });
      });
    });
  }

  // ---------- 按鈕 ----------
  function doInsert() {
    var latex = currentLatex();
    if (!latex.replace(/\s/g, '')) { setStatus('請先輸入方程式', 'err'); return; }
    if (HOST === 'Excel') { doExcelCopy(); return; }
    if (HOST !== 'Word') {
      setStatus('目前不是在 Word 裡開啟，只能複製。請在 Word 或 Excel 的增益集中使用。', 'err');
      return;
    }
    $('btnInsert').disabled = true;
    setStatus('插入中…');
    wordInsert(latex).then(function (fallback) {
      pushHistory(latex);
      setStatus(fallback ? '已插入（用備用方式，後面若多了換行請按 Delete 接回）' : '已插入', 'ok');
    })['catch'](function (e) {
      setStatus('插入失敗：' + errText(e), 'err');
    }).then(function () { $('btnInsert').disabled = false; });
  }

  function doLoad() {
    if (HOST !== 'Word') { setStatus('讀取方程式只能在 Word 使用', 'err'); return; }
    setStatus('讀取中…');
    wordLoad().then(function (n) {
      setStatus('已載入' + (n > 1 ? '第 1 個（共 ' + n + ' 個）' : '') + '。修改後按「插入」會取代選取的方程式。', 'ok');
      if (mf && !texMode) mf.focus();
    })['catch'](function (e) { setStatus(errText(e), 'err'); });
  }

  function doFix(undo) {
    setStatus('處理中…');
    wordFixFractions(undo).then(function (n) {
      if (!n) setStatus('選取範圍裡沒有行內方程式的分數', 'err');
      else setStatus((undo ? '已還原 ' : '已放大 ') + n + ' 個分數（可按 Ctrl+Z 復原）', 'ok');
    })['catch'](function (e) { setStatus('處理失敗：' + errText(e), 'err'); });
  }

  function doExcelCopy() {
    var latex = currentLatex();
    if (!latex.replace(/\s/g, '')) { setStatus('請先輸入方程式', 'err'); return; }
    if (copyText(Latex2Omml.toUnicodeMath(latex))) {
      pushHistory(latex);
      setStatus('已複製。到 Excel 按 Alt+= → Ctrl+V → 按「專業」。', 'ok');
    } else setStatus('無法自動複製，請到「LaTeX 原始碼」手動複製。', 'err');
  }

  // ---------- 啟動 ----------
  function initEditor() {
    if (window.MathQuill) {
      mq = MathQuill.getInterface(2);
      mf = mq.MathField($('mathField'), fieldConfig({
        edit: onMathEdit,
        enter: doInsert,
        moveOutOf: function (dir) { if (dir === mq.R && grid) focusCell(0, 0); },
        downOutOf: function () { if (grid) focusCell(0, 0); }
      }));
      active = mf;
      $('mathField').addEventListener('focusin', function () { active = mf; });

      $('gridAddRow').onclick = function () { addRow(grid.r + 1); };
      $('gridDelRow').onclick = function () { delRow(grid.r); };
      $('gridAddCol').onclick = function () { addCol(grid.c + 1); };
      $('gridDelCol').onclick = function () { delCol(grid.c); };
      $('gridRemove').onclick = function () { removeGrid(); onMathEdit(); mf.focus(); };
    } else {
      // CDN 載不到時，至少還能用 LaTeX 輸入
      $('mathField').textContent = '（視覺編輯器載入失敗，請用下方 LaTeX 輸入）';
      setTexMode(true);
    }

    var tabs = document.querySelectorAll('.tab');
    for (var i = 0; i < tabs.length; i++) {
      tabs[i].onclick = function () { buildPalette(+this.getAttribute('data-tab')); };
    }
    buildPalette(0);

    var opts = load(OPTIONS_KEY, null);
    if (opts) {
      $('optNoShrink').checked = opts.noShrink !== false;
      $('optDisplay').checked = !!opts.display;
      $('optItalicNum').checked = !!opts.italicNum;
    }
    // 編輯框也跟著顯示斜體數字（MathQuill 自己有 class，不能整個覆蓋）
    var showItalic = function () {
      if ($('optItalicNum').checked) $('editorBox').classList.add('italic-num');
      else $('editorBox').classList.remove('italic-num');
    };
    $('optNoShrink').onchange = $('optDisplay').onchange = $('optItalicNum').onchange = function () {
      store(OPTIONS_KEY, options());
      showItalic();
    };
    showItalic();

    var ta = $('latex');
    ta.oninput = onLatexInput;
    ta.onkeyup = onLatexInput; // IE11 的 textarea 有時不觸發 input
    ta.onkeydown = function (e) {
      if (e.keyCode === 13 && (e.ctrlKey || e.metaKey)) { e.preventDefault(); doInsert(); }
    };

    $('btnInsert').onclick = doInsert;
    $('btnLoad').onclick = doLoad;
    $('btnClear').onclick = function () { setLatex(''); setStatus(''); if (mf) mf.focus(); };
    $('btnCopyTex').onclick = function () {
      var ok = copyText(currentLatex());
      setStatus(ok ? '已複製 LaTeX' : '複製失敗，請手動選取複製', ok ? 'ok' : 'err');
    };
    $('btnCopyUm').onclick = function () {
      var ok = copyText(Latex2Omml.toUnicodeMath(currentLatex()));
      setStatus(ok ? '已複製線性格式（貼到方程式後按「專業」）' : '複製失敗', ok ? 'ok' : 'err');
    };
    $('btnFix').onclick = function () { doFix(false); };
    $('btnUnfix').onclick = function () { doFix(true); };
    $('btnExcelCopy').onclick = doExcelCopy;
    renderHistory();
    $('version').textContent = '版本 ' + VERSION;
    if (mf) mf.focus();
  }

  function onHost(host) {
    HOST = host;
    if (host === 'Word') {
      $('wordTools').className = 'section';
    } else if (host === 'Excel') {
      $('excelTools').className = 'section';
      $('btnInsert').textContent = '複製到 Excel';
      $('btnLoad').className = 'hidden';
      $('optDisplay').parentNode.className = 'hidden';
      $('optNoShrink').parentNode.className = 'hidden';
    } else {
      setStatus('提示：在 Word 或 Excel 的增益集裡開啟才能插入；這裡可以先試用編輯和複製。');
    }
  }

  initEditor();
  if (window.Office && Office.onReady) {
    Office.onReady(function (info) {
      var h = info && info.host;
      onHost(h === Office.HostType.Word ? 'Word' : h === Office.HostType.Excel ? 'Excel' : null);
    });
  } else onHost(null);
}());
