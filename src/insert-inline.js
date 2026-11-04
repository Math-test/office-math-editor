/*
 * 在段落中間插入「行內」方程式而不切斷段落。
 *
 * Word 的 insertOoxml 插入 <w:p> 時一定會多一個段落記號，句子會被切成兩行。
 * 做法：先在游標處插入佔位文字（insertText 不會切段落）→ 讀整段 OOXML →
 *      把佔位文字所在的 run 拆開、中間放 m:oMath → 整段寫回。
 * 純 ES5 + 標準 DOM API。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./latex2omml.js'));
  else root.InsertInline = factory(root.Latex2Omml);
}(this, function (Latex2Omml) {
  'use strict';

  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  var XML_NS = 'http://www.w3.org/XML/1998/namespace';
  // 純英數：特殊符號會被 Word 換字型而拆成好幾個 run，就找不到了
  var MARKER = 'EQMARK7731';

  function isW(n, name) { return n && n.nodeType === 1 && n.namespaceURI === W && (n.localName || n.nodeName.replace(/^.*:/, '')) === name; }

  function omathElement(doc, latex, opt) {
    // 用暫存文件解析 OMML 字串，再匯入目標文件
    var xml = '<w:p ' + Latex2Omml.NAMESPACES + '>' + Latex2Omml.toOmml(latex, opt) + '</w:p>';
    var tmp = opt.parseXml(xml);
    return doc.importNode(tmp.documentElement.firstChild, true);
  }

  function runSizePt(run) {
    for (var c = run.firstChild; c; c = c.nextSibling) {
      if (!isW(c, 'rPr')) continue;
      for (var s = c.firstChild; s; s = s.nextSibling) {
        if (isW(s, 'sz')) {
          var v = parseInt(s.getAttributeNS(W, 'val') || s.getAttribute('w:val'), 10);
          if (v) return v / 2;
        }
      }
    }
    return 0;
  }

  function setText(t, text) {
    while (t.firstChild) t.removeChild(t.firstChild);
    t.appendChild(t.ownerDocument.createTextNode(text));
    t.setAttributeNS(XML_NS, 'xml:space', 'preserve');
  }

  /**
   * 把文件中的佔位文字換成方程式
   * @param {Document} doc 段落的 OOXML（Flat OPC）
   * @param {string} marker 佔位文字
   * @param {string} latex
   * @param {object} opt { fontSize, noShrink, parseXml(string)→Document, cursorMarker（選用）}
   * @returns {boolean} 是否有找到並替換
   */
  function replaceMarker(doc, marker, latex, opt) {
    opt = opt || {};
    if (!opt.parseXml) {
      opt.parseXml = function (s) { return new DOMParser().parseFromString(s, 'text/xml'); };
    }
    var ts = doc.getElementsByTagNameNS(W, 't');
    for (var i = 0; i < ts.length; i++) {
      var t = ts[i], text = t.textContent, at = text.indexOf(marker);
      if (at < 0) continue;
      var run = t.parentNode;
      if (!isW(run, 'r')) continue;
      var size = runSizePt(run) || opt.fontSize || 12;
      var math = omathElement(doc, latex, { fontSize: size, noShrink: opt.noShrink, parseXml: opt.parseXml });
      var before = text.slice(0, at), after = text.slice(at + marker.length);
      var parent = run.parentNode;
      var idx = childIndex(run, t);
      // 佔位之後的內容 → 複製 run（保留格式），只留 t 之後的子元素
      var tail = run.cloneNode(true);
      var tailKids = elementChildren(tail);
      for (var k = 0; k < idx; k++) if (!isW(tailKids[k], 'rPr')) tail.removeChild(tailKids[k]);
      setText(tailKids[idx], after);
      if (!after && !hasOtherContent(tail, tailKids[idx])) tail = null;
      // 原 run 只留 t 之前的子元素與佔位前的文字
      var sib = t.nextSibling;
      while (sib) { var nx = sib.nextSibling; run.removeChild(sib); sib = nx; }
      setText(t, before);
      parent.insertBefore(math, run.nextSibling);
      var anchor = math;
      if (opt.cursorMarker) {
        // 方程式後面放一段標記文字，寫回 Word 後搜尋它來放游標（之後刪掉）
        var cur = run.cloneNode(false);
        var rpr = firstRPr(run);
        if (rpr) cur.appendChild(rpr.cloneNode(true));
        var ct = doc.createElementNS(W, 'w:t');
        setText(ct, opt.cursorMarker);
        cur.appendChild(ct);
        parent.insertBefore(cur, math.nextSibling);
        anchor = cur;
      }
      if (tail) parent.insertBefore(tail, anchor.nextSibling);
      if (!before && !hasOtherContent(run, t)) parent.removeChild(run);
      return true;
    }
    return false;
  }

  function firstRPr(run) {
    for (var c = run.firstChild; c; c = c.nextSibling) if (isW(c, 'rPr')) return c;
    return null;
  }
  function elementChildren(n) {
    var out = [];
    for (var c = n.firstChild; c; c = c.nextSibling) if (c.nodeType === 1) out.push(c);
    return out;
  }
  function childIndex(parent, el) {
    var kids = elementChildren(parent);
    for (var i = 0; i < kids.length; i++) if (kids[i] === el) return i;
    return -1;
  }
  function hasOtherContent(run, t) {
    for (var c = run.firstChild; c; c = c.nextSibling) {
      if (c === t || isW(c, 'rPr')) continue;
      return true;
    }
    return false;
  }

  return { replaceMarker: replaceMarker, MARKER: MARKER };
}));
