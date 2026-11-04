/*
 * 把 Word 文件（OOXML）裡「行內方程式」的分數放大，抵銷 Word 的自動縮小。
 * 用在：選取範圍 getOoxml() → 修改 → insertOoxml() 取代。
 * 純 ES5 + 標準 DOM API。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FixFrac = factory();
}(this, function () {
  'use strict';

  var M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
  var W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  var SCALE = 1.4; // 跟 latex2omml.js 的 INLINE_FRACTION_SCALE 一樣
  // 這些容器裡的分數本來就該小（上下標、根指數、極限下方、已在分數裡）
  var SKIP_PARENTS = { num: 1, den: 1, sub: 1, sup: 1, deg: 1, lim: 1 };

  function local(n) { return n.localName || String(n.nodeName).replace(/^.*:/, ''); }
  function isM(n, name) { return n && n.nodeType === 1 && n.namespaceURI === M && local(n) === name; }
  function isW(n, name) { return n && n.nodeType === 1 && n.namespaceURI === W && local(n) === name; }
  function firstChild(n, test) {
    for (var c = n.firstChild; c; c = c.nextSibling) if (test(c)) return c;
    return null;
  }
  function wAttr(el, name) {
    var v = el.getAttributeNS ? el.getAttributeNS(W, name) : null;
    return v || el.getAttribute('w:' + name);
  }

  // 方程式所在段落的文字字級（半點）：先看前後一般文字 run，再看段落記號
  function baseSize(oMath, fallbackHalfPts) {
    var p = oMath.parentNode;
    while (p && !isW(p, 'p')) p = p.parentNode;
    if (!p) return fallbackHalfPts;
    var runs = p.getElementsByTagNameNS(W, 'r');
    for (var i = 0; i < runs.length; i++) {
      var r = runs[i], anc = r.parentNode, inMath = false;
      while (anc && anc !== p) { if (anc.namespaceURI === M) { inMath = true; break; } anc = anc.parentNode; }
      if (inMath) continue;
      var sz = szOf(firstChild(r, function (c) { return isW(c, 'rPr'); }));
      if (sz) return sz;
    }
    var ppr = firstChild(p, function (c) { return isW(c, 'pPr'); });
    var prpr = ppr && firstChild(ppr, function (c) { return isW(c, 'rPr'); });
    return szOf(prpr) || fallbackHalfPts;
  }
  function szOf(rpr) {
    if (!rpr) return 0;
    var sz = firstChild(rpr, function (c) { return isW(c, 'sz'); });
    return sz ? parseInt(wAttr(sz, 'val'), 10) || 0 : 0;
  }

  function topLevelFractions(oMath) {
    var all = oMath.getElementsByTagNameNS(M, 'f'), out = [];
    for (var i = 0; i < all.length; i++) {
      var p = all[i].parentNode, skip = false;
      while (p && p !== oMath) {
        if (p.namespaceURI === M && SKIP_PARENTS[local(p)]) { skip = true; break; }
        p = p.parentNode;
      }
      if (!skip) out.push(all[i]);
    }
    return out;
  }

  // 在 m:r 設定 w:sz / w:szCs（size=0 表示移除）
  function setRunSize(doc, run, size) {
    var rpr = firstChild(run, function (c) { return isW(c, 'rPr'); });
    if (!rpr) {
      if (!size) return;
      rpr = doc.createElementNS(W, 'w:rPr');
      var mt = firstChild(run, function (c) { return isM(c, 't'); });
      run.insertBefore(rpr, firstChild(run, function (c) { return !isM(c, 'rPr'); }) || mt);
    }
    ['sz', 'szCs'].forEach(function (name) {
      var el = firstChild(rpr, function (c) { return isW(c, name); });
      if (!size) { if (el) rpr.removeChild(el); return; }
      if (!el) { el = doc.createElementNS(W, 'w:' + name); rpr.appendChild(el); }
      el.setAttributeNS(W, 'w:val', String(size));
    });
  }

  /**
   * @param {Document} doc 已解析的 OOXML（Flat OPC 或 document.xml）
   * @param {object} opt  { fontSize: 選取範圍字級 pt（找不到段落字級時用）, undo: true 表示還原成 Word 預設大小 }
   * @returns {number} 處理的分數數量
   */
  function apply(doc, opt) {
    opt = opt || {};
    var fallback = Math.round((opt.fontSize || 12) * 2);
    var maths = doc.getElementsByTagNameNS(M, 'oMath'), count = 0;
    for (var i = 0; i < maths.length; i++) {
      var om = maths[i];
      if (isM(om.parentNode, 'oMathPara')) continue; // 獨立方程式不會縮小
      var target = opt.undo ? 0 : Math.round(baseSize(om, fallback) * SCALE);
      var fracs = topLevelFractions(om);
      for (var j = 0; j < fracs.length; j++) {
        var parts = [firstChild(fracs[j], function (c) { return isM(c, 'num'); }),
                     firstChild(fracs[j], function (c) { return isM(c, 'den'); })];
        for (var k = 0; k < parts.length; k++) {
          if (!parts[k]) continue;
          var runs = parts[k].getElementsByTagNameNS(M, 'r');
          for (var r = 0; r < runs.length; r++) setRunSize(doc, runs[r], target);
          // 結構元素（根號、括號…）的 ctrlPr 也要一起，否則符號大小不一致
          var ctrls = parts[k].getElementsByTagNameNS(M, 'ctrlPr');
          for (var c = 0; c < ctrls.length; c++) {
            var rpr = firstChild(ctrls[c], function (x) { return isW(x, 'rPr'); });
            if (!rpr && target) { rpr = doc.createElementNS(W, 'w:rPr'); ctrls[c].appendChild(rpr); }
            if (rpr) setRunSize(doc, ctrls[c], target);
          }
        }
        count++;
      }
    }
    return count;
  }

  return { apply: apply, SCALE: SCALE };
}));
