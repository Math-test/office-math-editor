/*
 * OMML（Word 方程式）→ LaTeX
 *
 * 讓 Word 裡既有的方程式（不論是這個工具插入的、或用 Word 內建方程式編輯器打的）
 * 都能載回編輯器修改。純 ES5，只用標準 DOM API（瀏覽器 DOMParser / Node 的 xmldom 皆可）。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Omml2Latex = factory();
}(this, function () {
  'use strict';

  var M_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';

  // Unicode → LaTeX 指令（反查 latex2omml 的符號表，挑常用寫法）
  var CHAR2TEX = {
    'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', 'δ': '\\delta', 'ϵ': '\\epsilon', 'ε': '\\varepsilon',
    'ζ': '\\zeta', 'η': '\\eta', 'θ': '\\theta', 'ϑ': '\\vartheta', 'ι': '\\iota', 'κ': '\\kappa',
    'λ': '\\lambda', 'μ': '\\mu', 'ν': '\\nu', 'ξ': '\\xi', 'π': '\\pi', 'ρ': '\\rho', 'σ': '\\sigma',
    'τ': '\\tau', 'υ': '\\upsilon', 'ϕ': '\\phi', 'φ': '\\varphi', 'χ': '\\chi', 'ψ': '\\psi', 'ω': '\\omega',
    'Γ': '\\Gamma', 'Δ': '\\Delta', 'Θ': '\\Theta', 'Λ': '\\Lambda', 'Ξ': '\\Xi', 'Π': '\\Pi', 'Σ': '\\Sigma',
    'Φ': '\\Phi', 'Ψ': '\\Psi', 'Ω': '\\Omega',
    '×': '\\times', '÷': '\\div', '⋅': '\\cdot', '·': '\\cdot', '±': '\\pm', '∓': '\\mp', '∘': '\\circ',
    '∪': '\\cup', '∩': '\\cap', '∖': '\\setminus', '∧': '\\wedge', '∨': '\\vee',
    '≤': '\\le', '≥': '\\ge', '≠': '\\ne', '≈': '\\approx', '∼': '\\sim', '≅': '\\cong', '≡': '\\equiv',
    '∝': '\\propto', '∈': '\\in', '∉': '\\notin', '⊂': '\\subset', '⊃': '\\supset', '⊆': '\\subseteq',
    '⊇': '\\supseteq', '⊥': '\\perp', '∥': '\\parallel', '→': '\\to', '←': '\\leftarrow',
    '↔': '\\leftrightarrow', '⇒': '\\Rightarrow', '⇐': '\\Leftarrow', '⇔': '\\Leftrightarrow',
    '∞': '\\infty', '∂': '\\partial', '∇': '\\nabla', '∀': '\\forall', '∃': '\\exists', '¬': '\\neg',
    '∅': '\\emptyset', '∠': '\\angle', '△': '\\triangle', '°': '^\\circ', '′': "'", '∴': '\\therefore',
    '∵': '\\because', '⋯': '\\cdots', '…': '\\ldots', '−': '-', 'ℝ': '\\mathbb{R}', 'ℕ': '\\mathbb{N}',
    'ℤ': '\\mathbb{Z}', 'ℚ': '\\mathbb{Q}', 'ℂ': '\\mathbb{C}', '%': '\\%', '{': '\\{', '}': '\\}',
    '#': '\\#', '&': '\\&', ' ': '\\,', ' ': '\\;', ' ': '\\quad ', ' ': '\\ ',
    '⟨': '\\langle', '⟩': '\\rangle'
  };
  var FUNCS = ['sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan', 'sinh', 'cosh',
    'tanh', 'log', 'ln', 'lg', 'exp', 'det', 'dim', 'gcd', 'max', 'min', 'arg', 'lim'];
  var NARY = { '∑': '\\sum', '∏': '\\prod', '∐': '\\coprod', '∫': '\\int', '∬': '\\iint', '∭': '\\iiint',
    '∮': '\\oint', '⋃': '\\bigcup', '⋂': '\\bigcap' };
  var ACC = { '⃗': '\\vec', '̂': '\\hat', '̃': '\\tilde', '̇': '\\dot', '̈': '\\ddot',
    '̅': '\\bar', '́': '\\acute', '̀': '\\grave', '̌': '\\check', '̆': '\\breve',
    '⃖': '\\overleftarrow' };
  var DELIM = { '(': '(', ')': ')', '[': '[', ']': ']', '{': '\\{', '}': '\\}', '|': '|', '‖': '\\|',
    '⟨': '\\langle', '⟩': '\\rangle', '⌊': '\\lfloor', '⌋': '\\rfloor', '⌈': '\\lceil', '⌉': '\\rceil', '': '.' };

  function isCJK(s) { return /[⺀-鿿豈-﫿＀-￯　-〿]/.test(s); }

  function local(n) { return n.localName || String(n.nodeName).replace(/^.*:/, ''); }
  function isM(n, name) { return n && n.nodeType === 1 && n.namespaceURI === M_NS && local(n) === name; }
  function kids(n) {
    var out = [], c = n ? n.childNodes : [];
    for (var i = 0; i < c.length; i++) if (c[i].nodeType === 1) out.push(c[i]);
    return out;
  }
  function childM(n, name) {
    var k = kids(n);
    for (var i = 0; i < k.length; i++) if (isM(k[i], name)) return k[i];
    return null;
  }
  // 讀 <m:xxxPr><m:prop m:val="..."/></m:xxxPr>
  function prop(n, prName, propName) {
    var pr = childM(n, prName);
    var p = pr && childM(pr, propName);
    if (!p) return null;
    var v = p.getAttributeNS ? p.getAttributeNS(M_NS, 'val') : null;
    if (v === null || v === '') v = p.getAttribute('m:val');
    return v === null ? '' : v;
  }
  function on(v) { return v === '' || v === '1' || v === 'on' || v === 'true'; }

  function group(s) { return '{' + s + '}'; }

  // 把 LaTeX 片段接起來：指令後面若接英文字母要補空白
  function join(parts) {
    var out = '';
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i];
      if (!p) continue;
      if (/\\[a-zA-Z]+$/.test(out) && /^[a-zA-Z0-9]/.test(p)) out += ' ';
      out += p;
    }
    return out;
  }

  function children(n) { return join(kids(n).map(conv)); }

  function runText(r) {
    var t = '', k = kids(r);
    for (var i = 0; i < k.length; i++) {
      if (local(k[i]) === 't') t += k[i].textContent;
      else if (local(k[i]) === 'tab') t += ' ';
    }
    return t;
  }

  function convRun(r) {
    var text = runText(r);
    if (!text) return '';
    var rpr = childM(r, 'rPr');
    var nor = rpr && childM(rpr, 'nor');
    var sty = rpr && childM(rpr, 'sty');
    if (nor) {
      if (isCJK(text) && !/[a-zA-Z]/.test(text)) return text;
      if (/^[0-9.]+$/.test(text)) return text; // 「數字也斜體」產生的數字
      if (/^[    ]+$/.test(text)) {
        return join(text.split('').map(function (c) { return c === ' ' ? '\\ ' : mapChar(c); }));
      }
      return '\\text{' + text + '}';
    }
    if (sty && /^[a-zA-Z]{2,}$/.test(text)) {
      return FUNCS.indexOf(text) >= 0 ? '\\' + text : '\\operatorname{' + text + '}';
    }
    var parts = [];
    for (var i = 0; i < text.length; i++) parts.push(mapChar(text.charAt(i)));
    return join(parts);
  }

  function mapChar(ch) {
    if (CHAR2TEX[ch] !== undefined) return CHAR2TEX[ch];
    return ch;
  }

  function delimTex(ch, side) {
    var d = DELIM[ch];
    return '\\' + side + (d === undefined ? ch : d);
  }

  function conv(n) {
    if (n.nodeType !== 1) return '';
    if (n.namespaceURI !== M_NS) return ''; // w:rPr 等忽略
    var name = local(n), base, sub, sup;
    switch (name) {
      case 'oMathPara':
        return join(kids(n).filter(function (k) { return isM(k, 'oMath'); }).map(conv));
      case 'oMath': case 'e': case 'num': case 'den': case 'deg': case 'lim': case 'sub': case 'sup':
      case 'fName': case 'box': case 'phant':
        return children(n);
      case 'r':
        return convRun(n);
      case 'f': {
        var type = prop(n, 'fPr', 'type');
        var num = children(childM(n, 'num')), den = children(childM(n, 'den'));
        if (type === 'noBar') return '\\genfrac{}{}{0pt}{}' + group(num) + group(den);
        if (type === 'lin') return num + '/' + den;
        return '\\frac' + group(num) + group(den);
      }
      case 'rad': {
        var degHide = prop(n, 'radPr', 'degHide');
        var body = children(childM(n, 'e'));
        var deg = children(childM(n, 'deg'));
        if ((degHide !== null && on(degHide)) || !deg) return '\\sqrt' + group(body);
        return '\\sqrt[' + deg + ']' + group(body);
      }
      case 'sSup':
        return join([conv(childM(n, 'e')), '^' + group(children(childM(n, 'sup')))]);
      case 'sSub':
        return join([conv(childM(n, 'e')), '_' + group(children(childM(n, 'sub')))]);
      case 'sSubSup':
        return join([conv(childM(n, 'e')), '_' + group(children(childM(n, 'sub'))), '^' + group(children(childM(n, 'sup')))]);
      case 'sPre':
        return join(['{}_' + group(children(childM(n, 'sub'))) + '^' + group(children(childM(n, 'sup'))), conv(childM(n, 'e'))]);
      case 'd': {
        var beg = prop(n, 'dPr', 'begChr'), end = prop(n, 'dPr', 'endChr'), sep = prop(n, 'dPr', 'sepChr');
        if (beg === null) beg = '(';
        if (end === null) end = ')';
        if (sep === null) sep = '|';
        var es = kids(n).filter(function (k) { return isM(k, 'e'); });
        // {  + eqArr → cases；( + 無分數線分數 → binom；括號包矩陣 → pmatrix 等
        if (es.length === 1) {
          var only = kids(es[0]);
          if (only.length === 1 && isM(only[0], 'eqArr') && beg === '{' && end === '') {
            return '\\begin{cases}' + eqArrRows(only[0], ' & ') + '\\end{cases}';
          }
          if (only.length === 1 && isM(only[0], 'f') && prop(only[0], 'fPr', 'type') === 'noBar' && beg === '(' && end === ')') {
            return '\\binom' + group(children(childM(only[0], 'num'))) + group(children(childM(only[0], 'den')));
          }
          if (only.length === 1 && isM(only[0], 'm')) {
            var env = { '()': 'pmatrix', '[]': 'bmatrix', '{}': 'Bmatrix', '||': 'vmatrix', '‖‖': 'Vmatrix' }[beg + end];
            if (env) return matrix(only[0], env);
          }
        }
        var inner = es.map(conv).join(sep === '|' ? '\\mid ' : mapChar(sep));
        return delimTex(beg, 'left') + inner + delimTex(end, 'right');
      }
      case 'nary': {
        var chr = prop(n, 'naryPr', 'chr');
        if (chr === null || chr === '') chr = '∫';
        var cmd = NARY[chr] || chr;
        var subHide = prop(n, 'naryPr', 'subHide'), supHide = prop(n, 'naryPr', 'supHide');
        sub = (subHide !== null && on(subHide)) ? '' : children(childM(n, 'sub'));
        sup = (supHide !== null && on(supHide)) ? '' : children(childM(n, 'sup'));
        return join([cmd, sub ? '_' + group(sub) : '', sup ? '^' + group(sup) : '', ' ', children(childM(n, 'e'))]);
      }
      case 'func':
        return join([children(childM(n, 'fName')), ' ', children(childM(n, 'e'))]);
      case 'limLow': {
        base = children(childM(n, 'e'));
        var lim = children(childM(n, 'lim'));
        if (base === '\\lim' || base === 'lim' || base === '\\operatorname{lim}') return '\\lim_' + group(lim);
        return '\\underset' + group(lim) + group(base);
      }
      case 'limUpp':
        return '\\overset' + group(children(childM(n, 'lim'))) + group(children(childM(n, 'e')));
      case 'acc': {
        var ac = prop(n, 'accPr', 'chr');
        if (ac === null || ac === '') ac = '̂';
        var ab = children(childM(n, 'e'));
        if (ac === '⃗' && ab.replace(/\\[a-z]+/g, 'x').length > 1) return '\\overrightarrow' + group(ab);
        if (ac === '̂' && ab.length > 1) return '\\widehat' + group(ab);
        return (ACC[ac] || '\\hat') + group(ab);
      }
      case 'bar': {
        var pos = prop(n, 'barPr', 'pos');
        return (pos === 'top' ? '\\overline' : '\\underline') + group(children(childM(n, 'e')));
      }
      case 'groupChr': {
        var gc = prop(n, 'groupChrPr', 'chr');
        var gp = prop(n, 'groupChrPr', 'pos');
        var gb = children(childM(n, 'e'));
        if (gc === null || gc === '') gc = '⏟';
        if (gc === '⏜') return '\\overparen' + group(gb);
        if (gc === '⏝') return '\\underparen' + group(gb);
        if (gc === '⏞') return '\\overbrace' + group(gb);
        if (gc === '⏟') return '\\underbrace' + group(gb);
        return (gp === 'top' ? '\\overset' : '\\underset') + group(mapChar(gc)) + group(gb);
      }
      case 'borderBox':
        return '\\boxed' + group(children(childM(n, 'e')));
      case 'eqArr':
        return '\\begin{aligned}' + eqArrRows(n, '') + '\\end{aligned}';
      case 'm':
        return matrix(n, 'matrix');
      default:
        return children(n);
    }
  }

  function eqArrRows(n, cellSep) {
    return kids(n).filter(function (k) { return isM(k, 'e'); }).map(function (row) {
      var s = children(row);
      // latex2omml 用全形空白分隔 cases 的兩欄；aligned 裡的 & 是 Word 的對齊點
      if (cellSep) s = s.replace(/\\quad\s*/, cellSep);
      else s = s.replace(/\\&/g, '&');
      return s;
    }).join('\\\\');
  }

  function matrix(n, env) {
    var rows = kids(n).filter(function (k) { return isM(k, 'mr'); }).map(function (mr) {
      return kids(mr).filter(function (k) { return isM(k, 'e'); }).map(conv).join(' & ');
    });
    return '\\begin{' + env + '}' + rows.join('\\\\') + '\\end{' + env + '}';
  }

  /** OMML 元素（m:oMath / m:oMathPara）→ LaTeX */
  function fromElement(el) {
    return conv(el).replace(/\s+/g, ' ').replace(/^ | $/g, '');
  }

  /**
   * 從 Word getOoxml() 的結果找出方程式
   * @param {string} xml Flat OPC 或 document.xml
   * @param {function} parseXml (string) → Document（瀏覽器傳 DOMParser，Node 傳 xmldom）
   * @returns {{latex: string, display: boolean}[]}
   */
  function fromOoxml(xml, parseXml) {
    var doc = parseXml(xml);
    var out = [];
    // 依文件順序；oMathPara 裡的 oMath 標成 display
    var maths = doc.getElementsByTagNameNS(M_NS, 'oMath');
    for (var i = 0; i < maths.length; i++) {
      var p = maths[i].parentNode, nested = false;
      while (p) { if (isM(p, 'oMath')) { nested = true; break; } p = p.parentNode; }
      if (nested) continue;
      out.push({ latex: fromElement(maths[i]), display: isM(maths[i].parentNode, 'oMathPara') });
    }
    return out;
  }

  return { fromElement: fromElement, fromOoxml: fromOoxml, M_NS: M_NS };
}));
