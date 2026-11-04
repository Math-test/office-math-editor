/*
 * LaTeX → OMML（Word 方程式）轉換器
 *
 * 純 ES5，可在 Office 2016/2019（IE11 核心）、網頁版 Office、Node.js 執行。
 * 支援高中數學常用語法：分數、根號、上下標、括號、Σ/∫、lim、三角/對數函數、
 * 向量、線段、弧、矩陣、cases、希臘字母與常用符號。
 *
 * 「分數不縮小」：Word 會把行內方程式的分數縮成約 67%，
 * 這裡把最外層分數的分子、分母字級放大 1/0.67 倍抵銷，結果仍是可編輯的 Word 方程式。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Latex2Omml = factory();
}(this, function () {
  'use strict';

  // 實測：12pt 文字，分子分母設 17pt 時數字跟旁邊一樣高（Word 縮小的比例不是固定 0.67，1.49 會偏大）
  var INLINE_FRACTION_SCALE = 1.4;

  // ---------- 符號表 ----------
  var SYMBOLS = {
    // 希臘字母
    alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ', varepsilon: 'ε', zeta: 'ζ', eta: 'η',
    theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ', lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π',
    varpi: 'ϖ', rho: 'ρ', varrho: 'ϱ', sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ',
    varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
    Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π', Sigma: 'Σ', Upsilon: 'Υ',
    Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
    // 運算
    times: '×', div: '÷', cdot: '⋅', pm: '±', mp: '∓', ast: '∗', star: '⋆', circ: '∘', bullet: '∙',
    cup: '∪', cap: '∩', setminus: '∖', oplus: '⊕', otimes: '⊗', wedge: '∧', land: '∧', vee: '∨', lor: '∨',
    // 關係
    le: '≤', leq: '≤', ge: '≥', geq: '≥', ne: '≠', neq: '≠', approx: '≈', sim: '∼', simeq: '≃',
    cong: '≅', equiv: '≡', propto: '∝', ll: '≪', gg: '≫', leqslant: '⩽', geqslant: '⩾',
    'in': '∈', notin: '∉', ni: '∋', subset: '⊂', supset: '⊃', subseteq: '⊆', supseteq: '⊇',
    subsetneq: '⊊', perp: '⊥', parallel: '⫽', sslash: '⫽', // 台灣課本的平行是斜的
    nparallel: '∦', mid: '∣', nmid: '∤',
    // 箭頭
    to: '→', rightarrow: '→', leftarrow: '←', gets: '←', leftrightarrow: '↔', Rightarrow: '⇒',
    Leftarrow: '⇐', Leftrightarrow: '⇔', implies: '⇒', iff: '⇔', mapsto: '↦', uparrow: '↑',
    downarrow: '↓', longrightarrow: '⟶', Longrightarrow: '⟹',
    // 其他
    infty: '∞', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', neg: '¬', lnot: '¬',
    emptyset: '∅', varnothing: '∅', angle: '∠', measuredangle: '∡', triangle: '△', square: '□',
    Box: '□', degree: '°', prime: '′', therefore: '∴', because: '∵', cdots: '⋯', ldots: '…',
    dots: '…', vdots: '⋮', ddots: '⋱', hbar: 'ℏ', ell: 'ℓ', aleph: 'ℵ',
    lbrace: '{', rbrace: '}', langle: '⟨', rangle: '⟩', lfloor: '⌊', rfloor: '⌋', lceil: '⌈',
    rceil: '⌉', vert: '|', Vert: '‖', backslash: '\\', '%': '%', '$': '$', '#': '#', '&': '&',
    '_': '_', '{': '{', '}': '}', '|': '‖'
  };
  var SPACES = { ',': ' ', ':': ' ', ';': ' ', ' ': ' ', quad: ' ', qquad: '  ', '!': '' };
  var FUNCS = ['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh',
    'tanh', 'log', 'ln', 'lg', 'exp', 'det', 'deg', 'dim', 'gcd', 'max', 'min', 'sup', 'inf', 'arg', 'ker', 'Pr'];
  var LIMFUNCS = { lim: 'lim', limsup: 'lim sup', liminf: 'lim inf' };
  var NARY = { sum: '∑', prod: '∏', coprod: '∐', int: '∫', iint: '∬', iiint: '∭', oint: '∮', bigcup: '⋃', bigcap: '⋂' };
  var ACCENTS = { vec: '⃗', overrightarrow: '⃗', hat: '̂', widehat: '̂', tilde: '̃',
    widetilde: '̃', dot: '̇', ddot: '̈', bar: '̅', acute: '́', grave: '̀',
    check: '̌', breve: '̆', overleftarrow: '⃖' };
  var GROUPCHR = { overparen: ['⏜', 'top'], overarc: ['⏜', 'top'], frown: ['⏜', 'top'], underparen: ['⏝', 'bot'],
    overbrace: ['⏞', 'top'], underbrace: ['⏟', 'bot'] };
  var MATRIX_DELIMS = { matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'], Bmatrix: ['{', '}'],
    vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'], array: ['', ''] };
  var RELATIONS = '=<>≤≥≠≈≡∼≃≅∝≪≫⩽⩾∈∉∋⊂⊃⊆⊇⊊⊥∥⫽→←↔⇒⇐⇔↦⟶⟹';
  var DOUBLESTRUCK = { R: 'ℝ', N: 'ℕ', Z: 'ℤ', Q: 'ℚ', C: 'ℂ', P: 'ℙ' };

  function isCJK(ch) {
    return /[⺀-鿿豈-﫿＀-￯　-〿]/.test(ch);
  }

  // ---------- 詞法 ----------
  function tokenize(src) {
    var toks = [], i = 0, ch, m;
    while (i < src.length) {
      ch = src.charAt(i);
      if (ch === '\\') {
        m = /^\\([a-zA-Z]+)\s*/.exec(src.slice(i));
        if (m) { toks.push({ t: 'cmd', v: m[1] }); i += m[0].length; continue; }
        toks.push({ t: 'cmd', v: src.charAt(i + 1) || '\\' });
        i += 2;
        continue;
      }
      if (ch === '%') { // 註解到行尾
        while (i < src.length && src.charAt(i) !== '\n') i++;
        continue;
      }
      if (/\s/.test(ch)) { i++; continue; }
      // surrogate pair（罕用字）
      var code = src.charCodeAt(i);
      if (code >= 0xD800 && code <= 0xDBFF && i + 1 < src.length) {
        toks.push({ t: 'char', v: src.substr(i, 2) }); i += 2; continue;
      }
      if ('{}^_&'.indexOf(ch) >= 0) toks.push({ t: ch });
      else if (ch === "'") toks.push({ t: 'char', v: '′' });
      else toks.push({ t: 'char', v: ch });
      i++;
    }
    return toks;
  }

  // ---------- 語法 ----------
  function Parser(src) {
    this.src = src;
    this.toks = tokenize(src);
    this.pos = 0;
  }
  Parser.prototype.peek = function () { return this.toks[this.pos]; };
  Parser.prototype.next = function () { return this.toks[this.pos++]; };
  Parser.prototype.isCmd = function (tok, name) { return tok && tok.t === 'cmd' && tok.v === name; };

  // 讀到 stop(tok) 為真為止；回傳 atom 陣列
  Parser.prototype.parseList = function (stop) {
    var list = [], tok, atom;
    while ((tok = this.peek())) {
      if (stop && stop(tok)) break;
      if (tok.t === '}') break; // 不成對的 } 交給上層
      if (tok.t === '^' || tok.t === '_') {
        this.attachScripts(list, null);
        continue;
      }
      atom = this.parseAtom();
      if (atom === null) continue;
      list.push(atom);
      this.attachScripts(list, atom);
    }
    return list;
  };

  // 處理 x^2、x_1、x_1^2 與 x'（′ 已轉成字元）
  Parser.prototype.attachScripts = function (list, atom) {
    var tok = this.peek();
    if (!tok || (tok.t !== '^' && tok.t !== '_')) return;
    var base = atom ? list.pop() : { t: 'row', items: [] };
    var sub = null, sup = null;
    while ((tok = this.peek()) && (tok.t === '^' || tok.t === '_')) {
      this.next();
      var arg = this.parseArg();
      if (tok.t === '^') sup = sup ? { t: 'row', items: [sup, arg] } : arg;
      else sub = sub ? { t: 'row', items: [sub, arg] } : arg;
    }
    // nary / lim 的上下限直接放進節點
    if (base.t === 'nary' || base.t === 'lim') {
      if (sub) base.sub = sub;
      if (sup) base.sup = sup;
      list.push(base);
      return;
    }
    if (base.t === 'func' && !base.sub && !base.sup) {
      base.sub = sub; base.sup = sup;
      list.push(base);
      return;
    }
    list.push({ t: 'scripts', base: base, sub: sub, sup: sup });
  };

  // 單一參數：{...} 或一個 token
  Parser.prototype.parseArg = function () {
    var tok = this.peek();
    if (!tok) return { t: 'row', items: [] };
    if (tok.t === '{') {
      this.next();
      var items = this.parseList();
      if (this.peek() && this.peek().t === '}') this.next();
      return items.length === 1 ? items[0] : { t: 'row', items: items };
    }
    // 沒有大括號時只取一個字元：x^23 是 x² 後面接 3、\log_2 8 是 log₂ 8
    if (tok.t === 'char') {
      this.next();
      return isCJK(tok.v) ? { t: 'text', v: tok.v } : { t: 'ord', v: tok.v, kind: /[0-9]/.test(tok.v) ? 'num' : undefined };
    }
    var a = this.parseAtom();
    return a || { t: 'row', items: [] };
  };

  Parser.prototype.parseOptArg = function () {
    var tok = this.peek();
    if (tok && tok.t === 'char' && tok.v === '[') {
      this.next();
      var items = this.parseList(function (t) { return t.t === 'char' && t.v === ']'; });
      if (this.peek()) this.next();
      return { t: 'row', items: items };
    }
    return null;
  };

  // 讀一個分隔符號：( ) [ ] \{ \} | \| . \langle 等
  Parser.prototype.parseDelim = function () {
    var tok = this.next();
    if (!tok) return '';
    if (tok.t === 'char') return tok.v === '.' ? '' : tok.v;
    if (tok.t === 'cmd') {
      if (tok.v === '{' || tok.v === 'lbrace') return '{';
      if (tok.v === '}' || tok.v === 'rbrace') return '}';
      if (tok.v === '|' || tok.v === 'Vert') return '‖';
      return SYMBOLS[tok.v] || '';
    }
    return '';
  };

  Parser.prototype.parseTextArg = function () {
    // \text{...}：原樣取字，不解析
    var tok = this.peek();
    if (!tok || tok.t !== '{') return '';
    var depth = 0, out = '';
    while ((tok = this.next())) {
      if (tok.t === '{') { if (depth++ > 0) out += '{'; continue; }
      if (tok.t === '}') { if (--depth === 0) break; out += '}'; continue; }
      if (tok.t === 'char') out += tok.v;
      else if (tok.t === 'cmd') out += (SPACES[tok.v] !== undefined ? ' ' : (SYMBOLS[tok.v] || ''));
      else out += tok.t;
    }
    return out;
  };

  Parser.prototype.parseEnv = function (name) {
    // \begin{name} ... \end{name}；以 & 分欄、\\ 分列
    var rows = [], row = [], self = this;
    var stop = function (t) { return t.t === '&' || self.isCmd(t, '\\') || self.isCmd(t, 'end'); };
    if (name === 'array') this.parseArg(); // 欄位格式 {cc} 略過
    for (;;) {
      var cell = this.parseList(stop);
      row.push({ t: 'row', items: cell });
      var tok = this.next();
      if (!tok) break;
      if (tok.t === '&') continue;
      if (this.isCmd(tok, '\\')) { rows.push(row); row = []; continue; }
      if (this.isCmd(tok, 'end')) { this.parseArg(); break; }
    }
    if (row.length && !(row.length === 1 && !row[0].items.length)) rows.push(row);
    return rows;
  };

  Parser.prototype.parseAtom = function () {
    var tok = this.next(), v;
    if (!tok) return null;
    if (tok.t === '{') {
      var items = this.parseList();
      if (this.peek() && this.peek().t === '}') this.next();
      return { t: 'row', items: items };
    }
    if (tok.t === '&') return { t: 'ord', v: ' ' };
    if (tok.t === 'char') {
      v = tok.v;
      if (/[0-9.]/.test(v)) {
        // 連續數字合成一個 run
        while (this.peek() && this.peek().t === 'char' && /[0-9.]/.test(this.peek().v)) v += this.next().v;
        return { t: 'ord', v: v, kind: 'num' };
      }
      if (isCJK(v)) {
        while (this.peek() && this.peek().t === 'char' && isCJK(this.peek().v)) v += this.next().v;
        return { t: 'text', v: v };
      }
      return { t: 'ord', v: v };
    }
    // ---- 指令 ----
    var c = tok.v;
    if (c === 'frac' || c === 'dfrac' || c === 'tfrac' || c === 'cfrac') {
      var num = this.parseArg(), den = this.parseArg();
      return { t: 'frac', num: num, den: den, display: c === 'dfrac' || c === 'cfrac' };
    }
    if (c === 'genfrac') {
      // \genfrac{左}{右}{線粗}{樣式}{分子}{分母}
      var gl = flatText(this.parseArg()), gr = flatText(this.parseArg());
      var thick = flatText(this.parseArg());
      this.parseArg();
      var gnum = this.parseArg(), gden = this.parseArg();
      var gf = { t: 'frac', num: gnum, den: gden, noBar: /^0/.test(thick) };
      return gl || gr ? { t: 'delim', open: gl, close: gr, body: gf } : gf;
    }
    if (c === 'binom' || c === 'dbinom' || c === 'tbinom') {
      var n = this.parseArg(), k = this.parseArg();
      return { t: 'delim', open: '(', close: ')', body: { t: 'frac', num: n, den: k, noBar: true } };
    }
    if (c === 'sqrt') {
      var deg = this.parseOptArg();
      return { t: 'sqrt', deg: deg, body: this.parseArg() };
    }
    if (c === 'left') {
      var open = this.parseDelim(), self = this;
      var body = this.parseList(function (t) { return self.isCmd(t, 'right'); });
      this.next(); // \right
      var close = this.parseDelim();
      return { t: 'delim', open: open, close: close, body: { t: 'row', items: body } };
    }
    if (c === 'right') return null; // 不成對
    if (c === 'big' || c === 'Big' || c === 'bigg' || c === 'Bigg' || c === 'bigl' || c === 'bigr' ||
        c === 'Bigl' || c === 'Bigr') return { t: 'ord', v: this.parseDelim() };
    if (c === 'text' || c === 'textrm' || c === 'mbox' || c === 'textnormal') {
      return { t: 'text', v: this.parseTextArg() };
    }
    if (c === 'mathrm' || c === 'operatorname' || c === 'rm' || c === 'mathup') {
      var a = this.parseArg();
      if (c === 'operatorname') return { t: 'func', name: flatText(a) };
      return { t: 'style', sty: 'p', body: a };
    }
    if (c === 'mathbb') {
      var bb = flatText(this.parseArg());
      return { t: 'ord', v: DOUBLESTRUCK[bb] || bb };
    }
    if (c === 'mathbf' || c === 'bf' || c === 'boldsymbol' || c === 'bm') return { t: 'style', sty: 'b', body: this.parseArg() };
    if (c === 'mathit') return { t: 'style', sty: 'i', body: this.parseArg() };
    if (c === 'displaystyle' || c === 'textstyle' || c === 'limits' || c === 'nolimits' || c === 'nonumber') return null;
    if (FUNCS.indexOf(c) >= 0) return { t: 'func', name: c };
    if (LIMFUNCS[c]) return { t: 'lim', name: LIMFUNCS[c] };
    if (NARY[c]) return { t: 'nary', chr: NARY[c] };
    if (ACCENTS[c]) return { t: 'acc', chr: ACCENTS[c], body: this.parseArg() };
    if (c === 'overline') return { t: 'bar', pos: 'top', body: this.parseArg() };
    if (c === 'underline') return { t: 'bar', pos: 'bot', body: this.parseArg() };
    if (GROUPCHR[c]) return { t: 'groupChr', chr: GROUPCHR[c][0], pos: GROUPCHR[c][1], body: this.parseArg() };
    if (c === 'overset' || c === 'stackrel') {
      var over = this.parseArg();
      return { t: 'limUpp', lim: over, base: this.parseArg() };
    }
    if (c === 'underset') {
      var under = this.parseArg();
      return { t: 'limLow', lim: under, base: this.parseArg() };
    }
    if (c === 'boxed' || c === 'fbox') return { t: 'box', body: this.parseArg() };
    if (c === 'begin') {
      var env = flatText(this.parseArg()).replace('*', '');
      var rows = this.parseEnv(env);
      if (env === 'cases') return { t: 'cases', rows: rows };
      if (env === 'aligned' || env === 'align' || env === 'gathered' || env === 'split') return { t: 'eqArr', rows: rows };
      var d = MATRIX_DELIMS[env] || ['', ''];
      return { t: 'matrix', rows: rows, open: d[0], close: d[1] };
    }
    if (c === 'end') { this.parseArg(); return null; }
    if (c === '\\' || c === 'newline') return { t: 'ord', v: ' ' };
    if (SPACES[c] !== undefined) return SPACES[c] ? { t: 'space', v: SPACES[c] } : null;
    if (c === 'hspace' || c === 'vspace') { this.parseArg(); return { t: 'space', v: ' ' }; }
    if (c === 'not') {
      var nt = this.parseAtom();
      return nt ? { t: 'ord', v: (nt.v || '') + '̸' } : null;
    }
    if (c === 'degree' || c === 'circ') return { t: 'ord', v: SYMBOLS[c] };
    if (SYMBOLS[c] !== undefined) return { t: 'ord', v: SYMBOLS[c] };
    // 未知指令：原樣顯示，至少不會整條失敗
    return { t: 'ord', v: '\\' + c };
  };

  function flatText(node) {
    if (!node) return '';
    if (node.t === 'row') return node.items.map(flatText).join('');
    if (node.v !== undefined) return node.v;
    return '';
  }

  // 後處理：func / lim / nary 吃掉後面的參數
  function isRelOrBin(atom) {
    return atom && atom.t === 'ord' && (RELATIONS.indexOf(atom.v) >= 0 || atom.v === '+' || atom.v === '-' ||
      atom.v === '−' || atom.v === ',' || atom.v === '±' || atom.v === '∓');
  }
  function isRel(atom) { return atom && atom.t === 'ord' && RELATIONS.indexOf(atom.v) >= 0; }

  function attachArgs(items) {
    var out = [], i, a, j;
    for (i = 0; i < items.length; i++) {
      a = fixNode(items[i]);
      if ((a.t === 'func' || a.t === 'lim') && !a.arg) {
        var nx = items[i + 1];
        if (nx && !isRelOrBin(nx) && !(nx.t === 'space')) {
          a.arg = fixNode(nx);
          i++;
        }
      } else if (a.t === 'nary' && !a.body) {
        // 積分：吃到 dx 為止；Σ/Π：吃到下一個 +、−、關係符號
        var body = [];
        var isInt = '∫∬∭∮'.indexOf(a.chr) >= 0;
        for (j = i + 1; j < items.length; j++) {
          var b = items[j];
          if (isRel(b)) break;
          if (!isInt && body.length && isRelOrBin(b)) break;
          body.push(b);
          if (isInt && b.t === 'ord' && b.v === 'd' && items[j + 1] && items[j + 1].t === 'ord' && /^[a-zA-Zα-ω]$/.test(items[j + 1].v)) {
            body.push(items[j + 1]);
            j += 2; // 跳過 d 和變數
            break;
          }
        }
        i = j - 1; // 迴圈的 i++ 會跳到 body 之後的第一個元素
        a.body = { t: 'row', items: attachArgs(body) };
      }
      out.push(a);
    }
    return out;
  }

  function fixNode(n) {
    if (!n || typeof n !== 'object') return n;
    if (n.t === 'row') return { t: 'row', items: attachArgs(n.items) };
    var k, c = {};
    for (k in n) {
      if (!Object.prototype.hasOwnProperty.call(n, k)) continue;
      var v = n[k];
      if (k === 'rows') {
        c.rows = v.map(function (r) { return r.map(fixNode); });
      } else if (v && typeof v === 'object' && v.t) c[k] = fixNode(v);
      else c[k] = v;
    }
    return c;
  }

  function parse(src) {
    var p = new Parser(src || '');
    var items = [];
    while (p.peek()) {
      items = items.concat(p.parseList());
      if (p.peek() && p.peek().t === '}') p.next(); // 多出的 }
    }
    return { t: 'row', items: attachArgs(items) };
  }

  // ---------- 產生 OMML ----------
  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  function run(text, ctx, opt) {
    opt = opt || {};
    var mrpr = '';
    if (opt.nor || opt.italic) mrpr = '<m:rPr><m:nor/></m:rPr>';
    else if (opt.sty || ctx.sty) mrpr = '<m:rPr><m:sty m:val="' + (opt.sty || ctx.sty) + '"/></m:rPr>';
    var wrpr = '';
    var fonts = opt.nor ? '<w:rFonts w:hint="eastAsia"/>' : '<w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/>';
    if (opt.italic) fonts += '<w:i/>';
    if (ctx.sz) wrpr = '<w:rPr>' + fonts + '<w:sz w:val="' + ctx.sz + '"/><w:szCs w:val="' + ctx.sz + '"/></w:rPr>';
    else wrpr = '<w:rPr>' + fonts + '</w:rPr>';
    return '<m:r>' + mrpr + wrpr + '<m:t xml:space="preserve">' + esc(text) + '</m:t></m:r>';
  }

  function ctrlPr(ctx) {
    return ctx.sz ? '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="' + ctx.sz + '"/><w:szCs w:val="' + ctx.sz + '"/></w:rPr></m:ctrlPr>'
      : '<m:ctrlPr><w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr></m:ctrlPr>';
  }

  function child(ctx, patch) {
    var c = {}, k;
    for (k in ctx) if (Object.prototype.hasOwnProperty.call(ctx, k)) c[k] = ctx[k];
    for (k in patch) if (Object.prototype.hasOwnProperty.call(patch, k)) c[k] = patch[k];
    return c;
  }

  function e(node, ctx, tag) {
    return '<m:' + (tag || 'e') + '>' + emit(node, ctx) + '</m:' + (tag || 'e') + '>';
  }

  function emit(n, ctx) {
    if (!n) return '';
    var i, out, sub;
    switch (n.t) {
      case 'row':
        out = '';
        for (i = 0; i < n.items.length; i++) out += emit(n.items[i], ctx);
        return out;
      case 'ord':
        if (n.v === '-') return run('−', ctx);
        // Word 的數字一律正體（Cambria Math 沒有斜體數字，sty=i 無效）；
        // 「數字也斜體」改用一般文字 + 斜體，實測可以
        if (n.kind === 'num' && ctx.italicNum && !ctx.sty) return run(n.v, ctx, { italic: true });
        return run(n.v, ctx);
      case 'text':
        return run(n.v, ctx, { nor: true });
      case 'space':
        return run(n.v, ctx, { nor: true });
      case 'style':
        return emit(n.body, child(ctx, { sty: n.sty }));
      case 'frac': {
        // 最外層、不在上下標裡的分數 → 分子分母放大（行內模式才需要）
        var fctx = ctx;
        if (ctx.noShrink && !ctx.inFrac && !ctx.inScript && ctx.baseSz) {
          fctx = child(ctx, { sz: Math.round(ctx.baseSz * INLINE_FRACTION_SCALE) });
        }
        fctx = child(fctx, { inFrac: true });
        var fpr = '<m:fPr>' + (n.noBar ? '<m:type m:val="noBar"/>' : '') + ctrlPr(ctx) + '</m:fPr>';
        return '<m:f>' + fpr + e(n.num, fctx, 'num') + e(n.den, fctx, 'den') + '</m:f>';
      }
      case 'sqrt':
        if (n.deg) {
          return '<m:rad><m:radPr>' + ctrlPr(ctx) + '</m:radPr>' + e(n.deg, child(ctx, { inScript: true }), 'deg') + e(n.body, ctx) + '</m:rad>';
        }
        return '<m:rad><m:radPr><m:degHide m:val="1"/>' + ctrlPr(ctx) + '</m:radPr><m:deg/>' + e(n.body, ctx) + '</m:rad>';
      case 'scripts': {
        sub = child(ctx, { inScript: true });
        if (n.sub && n.sup) {
          return '<m:sSubSup><m:sSubSupPr>' + ctrlPr(ctx) + '</m:sSubSupPr>' + e(n.base, ctx) + e(n.sub, sub, 'sub') + e(n.sup, sub, 'sup') + '</m:sSubSup>';
        }
        if (n.sup) return '<m:sSup><m:sSupPr>' + ctrlPr(ctx) + '</m:sSupPr>' + e(n.base, ctx) + e(n.sup, sub, 'sup') + '</m:sSup>';
        return '<m:sSub><m:sSubPr>' + ctrlPr(ctx) + '</m:sSubPr>' + e(n.base, ctx) + e(n.sub, sub, 'sub') + '</m:sSub>';
      }
      case 'delim': {
        var dpr = '<m:dPr><m:begChr m:val="' + esc(n.open) + '"/><m:endChr m:val="' + esc(n.close) + '"/>' + ctrlPr(ctx) + '</m:dPr>';
        return '<m:d>' + dpr + e(n.body, ctx) + '</m:d>';
      }
      case 'func': {
        var fname = run(n.name, ctx, { sty: 'p' });
        sub = child(ctx, { inScript: true });
        if (n.sub && n.sup) fname = '<m:sSubSup><m:sSubSupPr>' + ctrlPr(ctx) + '</m:sSubSupPr><m:e>' + fname + '</m:e>' + e(n.sub, sub, 'sub') + e(n.sup, sub, 'sup') + '</m:sSubSup>';
        else if (n.sub) fname = '<m:sSub><m:sSubPr>' + ctrlPr(ctx) + '</m:sSubPr><m:e>' + fname + '</m:e>' + e(n.sub, sub, 'sub') + '</m:sSub>';
        else if (n.sup) fname = '<m:sSup><m:sSupPr>' + ctrlPr(ctx) + '</m:sSupPr><m:e>' + fname + '</m:e>' + e(n.sup, sub, 'sup') + '</m:sSup>';
        if (!n.arg) return fname;
        return '<m:func><m:funcPr>' + ctrlPr(ctx) + '</m:funcPr><m:fName>' + fname + '</m:fName>' + e(n.arg, ctx) + '</m:func>';
      }
      case 'lim': {
        var lname = run(n.name, ctx, { sty: 'p' });
        if (n.sub) lname = '<m:limLow><m:limLowPr>' + ctrlPr(ctx) + '</m:limLowPr><m:e>' + lname + '</m:e>' + e(n.sub, child(ctx, { inScript: true }), 'lim') + '</m:limLow>';
        if (!n.arg) return lname;
        return '<m:func><m:funcPr>' + ctrlPr(ctx) + '</m:funcPr><m:fName>' + lname + '</m:fName>' + e(n.arg, ctx) + '</m:func>';
      }
      case 'nary': {
        sub = child(ctx, { inScript: true });
        var isInt = '∫∬∭∮'.indexOf(n.chr) >= 0;
        // 行內也讓 Σ 的上下限在符號上下方（像顯示模式），積分維持在右側
        var npr = '<m:naryPr><m:chr m:val="' + n.chr + '"/><m:limLoc m:val="' + (isInt ? 'subSup' : 'undOvr') + '"/>' +
          (n.sub ? '' : '<m:subHide m:val="1"/>') + (n.sup ? '' : '<m:supHide m:val="1"/>') + ctrlPr(ctx) + '</m:naryPr>';
        return '<m:nary>' + npr + e(n.sub, sub, 'sub') + e(n.sup, sub, 'sup') + e(n.body || { t: 'row', items: [] }, ctx) + '</m:nary>';
      }
      case 'acc':
        return '<m:acc><m:accPr><m:chr m:val="' + n.chr + '"/>' + ctrlPr(ctx) + '</m:accPr>' + e(n.body, ctx) + '</m:acc>';
      case 'bar':
        return '<m:bar><m:barPr><m:pos m:val="' + n.pos + '"/>' + ctrlPr(ctx) + '</m:barPr>' + e(n.body, ctx) + '</m:bar>';
      case 'groupChr':
        return '<m:groupChr><m:groupChrPr><m:chr m:val="' + n.chr + '"/><m:pos m:val="' + n.pos + '"/><m:vertJc m:val="' + (n.pos === 'top' ? 'bot' : 'top') + '"/>' + ctrlPr(ctx) + '</m:groupChrPr>' + e(n.body, ctx) + '</m:groupChr>';
      case 'limUpp':
        return '<m:limUpp><m:limUppPr>' + ctrlPr(ctx) + '</m:limUppPr>' + e(n.base, ctx) + e(n.lim, child(ctx, { inScript: true }), 'lim') + '</m:limUpp>';
      case 'limLow':
        return '<m:limLow><m:limLowPr>' + ctrlPr(ctx) + '</m:limLowPr>' + e(n.base, ctx) + e(n.lim, child(ctx, { inScript: true }), 'lim') + '</m:limLow>';
      case 'box':
        return '<m:borderBox><m:borderBoxPr>' + ctrlPr(ctx) + '</m:borderBoxPr>' + e(n.body, ctx) + '</m:borderBox>';
      case 'eqArr':
      case 'cases': {
        out = '<m:eqArr><m:eqArrPr>' + ctrlPr(ctx) + '</m:eqArrPr>';
        for (i = 0; i < n.rows.length; i++) {
          // cases 的「值 & 條件」中間空一格；aligned 的 & 交給 Word 對齊
          var cells = n.rows[i].map(function (c) { return emit(c, ctx); });
          out += '<m:e>' + cells.join(n.t === 'cases' ? run(' ', ctx, { nor: true }) : run('&', ctx)) + '</m:e>';
        }
        out += '</m:eqArr>';
        if (n.t === 'cases') {
          return '<m:d><m:dPr><m:begChr m:val="{"/><m:endChr m:val=""/>' + ctrlPr(ctx) + '</m:dPr><m:e>' + out + '</m:e></m:d>';
        }
        return out;
      }
      case 'matrix': {
        var cols = 0;
        for (i = 0; i < n.rows.length; i++) cols = Math.max(cols, n.rows[i].length);
        out = '<m:m><m:mPr><m:mcs><m:mc><m:mcPr><m:count m:val="' + cols + '"/><m:mcJc m:val="center"/></m:mcPr></m:mc></m:mcs>' + ctrlPr(ctx) + '</m:mPr>';
        for (i = 0; i < n.rows.length; i++) {
          out += '<m:mr>';
          for (var j = 0; j < cols; j++) out += e(n.rows[i][j] || { t: 'row', items: [] }, ctx);
          out += '</m:mr>';
        }
        out += '</m:m>';
        if (n.open || n.close) {
          return '<m:d><m:dPr><m:begChr m:val="' + esc(n.open) + '"/><m:endChr m:val="' + esc(n.close) + '"/>' + ctrlPr(ctx) + '</m:dPr><m:e>' + out + '</m:e></m:d>';
        }
        return out;
      }
      default:
        return '';
    }
  }

  /**
   * LaTeX → OMML
   * @param {string} latex
   * @param {object} [options]
   *   noShrink {boolean} 行內分數不縮小（預設 true）
   *   fontSize {number}  插入位置的字級（pt），用來計算放大後的分數字級，預設 12
   *   display  {boolean} 產生獨立一行的方程式（m:oMathPara）
   *   italicNum {boolean} 數字也用斜體（Word 預設只有字母斜體）
   * @returns {string} <m:oMath>…</m:oMath>（display 時外面包 m:oMathPara）
   */
  function toOmml(latex, options) {
    options = options || {};
    var display = !!options.display;
    var noShrink = options.noShrink !== false && !display;
    var ctx = { noShrink: noShrink, baseSz: Math.round((options.fontSize || 12) * 2), italicNum: !!options.italicNum };
    var body = emit(parse(latex), ctx);
    var math = '<m:oMath>' + body + '</m:oMath>';
    if (display) math = '<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>' + math + '</m:oMathPara>';
    return math;
  }

  // ---------- UnicodeMath（Office 線性格式）----------
  // Excel 增益集沒有插入方程式的 API；在 Excel 方程式裡貼上線性格式、按「專業」就會變成原生方程式。
  function um(n) {
    if (!n) return '';
    var i, out;
    switch (n.t) {
      case 'row':
        out = '';
        for (i = 0; i < n.items.length; i++) out += um(n.items[i]);
        return out;
      case 'ord': return n.v === '-' ? '−' : n.v;
      case 'text': return '"' + n.v + '"';
      case 'space': return n.v;
      case 'style': return um(n.body);
      case 'frac': // 結尾空格讓分母到此為止（否則後面的文字會被吃進分母）
        return n.noBar ? umGroup(n.num) + '¦' + umGroup(n.den) : umGroup(n.num) + '/' + umGroup(n.den) + ' ';
      case 'sqrt':
        return n.deg ? '√(' + um(n.deg) + '&' + um(n.body) + ')' : '√' + umGroup(n.body);
      case 'scripts':
        return umBase(n.base) + (n.sub ? '_' + umGroup(n.sub) : '') + (n.sup ? '^' + umGroup(n.sup) : '') + ' ';
      case 'delim':
        if (n.body && n.body.t === 'frac' && n.body.noBar) return '(' + um(n.body) + ')';
        return (n.open || '├') + um(n.body) + (n.close || '┤');
      // 函數、lim、Σ/∫ 的作用範圍用隱形括號 〖〗，一般括號會被當成看得見的括號
      case 'func':
        return n.name + (n.sub ? '_' + umGroup(n.sub) : '') + (n.sup ? '^' + umGroup(n.sup) : '') + '⁡' + (n.arg ? umInvisible(n.arg) : '');
      case 'lim':
        return n.name.replace(' ', ' ') + (n.sub ? '┬' + umGroup(n.sub) : '') + '⁡' + (n.arg ? umInvisible(n.arg) : '');
      case 'nary':
        return n.chr + (n.sub ? '_' + umGroup(n.sub) : '') + (n.sup ? '^' + umGroup(n.sup) : '') + '▒' + umInvisible(n.body || { t: 'row', items: [] });
      case 'acc': return umGroup(n.body) + n.chr;
      case 'bar': return (n.pos === 'top' ? '¯' : '▁') + umGroup(n.body);
      case 'groupChr': return n.chr + umGroup(n.body);
      case 'limUpp': return umGroup(n.base) + '┴' + umGroup(n.lim);
      case 'limLow': return umGroup(n.base) + '┬' + umGroup(n.lim);
      case 'box': return '▭' + umGroup(n.body);
      case 'cases': // 兩欄中間用全形空白（& 在 Word 是對齊點，不會留空）
        return '{█(' + umRows(n.rows, ' ') + ')┤';
      case 'eqArr':
        return '█(' + umRows(n.rows, '&') + ')';
      case 'matrix':
        return (n.open || '') + '■(' + umRows(n.rows, '&') + ')' + (n.close || '');
      default: return '';
    }
  }
  function umRows(rows, sep) {
    return rows.map(function (r) { return r.map(um).join(sep); }).join('@');
  }
  // 單一字元/數字不加括號，其他用 ( ) 包起來（Word 建構時會拿掉這層括號）
  function umGroup(n) {
    var s = um(n);
    if (/^[A-Za-z0-9.α-ωΑ-Ω∞]$/.test(s) || /^[0-9.]+$/.test(s)) return s;
    return '(' + s + ')';
  }
  function umInvisible(n) {
    var s = um(n);
    if (/^[A-Za-z0-9.α-ωΑ-Ω]$/.test(s) || /^[0-9.]+$/.test(s)) return s;
    return '〖' + s.replace(/ +$/, '') + '〗';
  }
  function umBase(n) {
    var s = um(n);
    if (n.t === 'delim' || /^[A-Za-z0-9α-ωΑ-Ω]$/.test(s) || /^[0-9.]+$/.test(s)) return s;
    return '〖' + s + '〗';
  }

  /** LaTeX → UnicodeMath 線性格式（貼到 Excel/Word 方程式後按「專業」） */
  function toUnicodeMath(latex) {
    return um(parse(latex)).replace(/ +$/, '');
  }

  var NS = 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
           'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  /** 包成 Word insertOoxml 可用的 Flat OPC 套件 */
  function toFlatOpc(latex, options) {
    var math = toOmml(latex, options);
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">' +
      '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships></pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">' +
      '<pkg:xmlData><w:document ' + NS + '><w:body><w:p>' + math + '</w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
      '</pkg:package>';
  }

  return {
    parse: parse,
    toOmml: toOmml,
    toFlatOpc: toFlatOpc,
    toUnicodeMath: toUnicodeMath,
    NAMESPACES: NS,
    INLINE_FRACTION_SCALE: INLINE_FRACTION_SCALE,
    SYMBOLS: SYMBOLS
  };
}));
