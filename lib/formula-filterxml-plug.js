/* formula-filterxml-plug.js — ★FILTERXML を エンジンに 繋ぐ★（2026-09-07）
 *
 *  ★XPath は ブラウザの 物を そのまま 使う★（自分で 作らない）
 *    ⇒ 画面では `window.DOMParser` / `window.XPathResult`
 *    ⇒ 試験では jsdom の 同じ 物を 入れる（★本番と 同じ 道具★）
 *
 *  見張り: tests/formula-filterxml.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaFilterXmlPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function つなぐ(H, F, 部品) {
    if (!H || !H.FunctionPlugin) return 0;
    if (部品) つなぐ.部品 = 部品;
    if (つなぐ.済み) return つなぐ.数;
    var CellError = H.CellError, ErrorType = H.ErrorType;
    var T = H.FunctionArgumentType || {};
    if (!CellError || !ErrorType || !T.ANY) return 0;

    function 値(自, ast, state, i) {
      if (!ast.args[i]) return null;
      var v = 自.evaluateAst(ast.args[i], state);
      if (v instanceof CellError) return v;
      return (v === null || v === undefined) ? null : v;
    }
    function 道具() {
      var d = つなぐ.部品;
      if (typeof d === 'function') d = d();
      if (d && d.DOMParser && d.XPathResult) return d;
      /* ★画面では そのまま 窓の 物★ */
      if (typeof self !== 'undefined' && self.DOMParser && self.XPathResult) return self;
      return null;
    }

    var Plug = class extends H.FunctionPlugin {
      filterxml(ast, state) {
        var x = 値(this, ast, state, 0);
        if (x instanceof CellError) return x;
        var p = 値(this, ast, state, 1);
        if (p instanceof CellError) return p;
        /* ★字で ない 物は そのまま 渡す★（lib が #VALUE! に する） */
        var 出 = F.取り出す(typeof x === 'string' ? x : x, typeof p === 'string' ? p : p, 道具());
        if (out誤り(出)) return new CellError(ErrorType[出.誤り] || ErrorType.VALUE);
        return H.SimpleRangeValue ? H.SimpleRangeValue.onlyValues(出) : 出;
      }
    };
    function out誤り(v) { return !!(v && v.誤り); }

    var ArraySize = H.ArraySize;
    Plug.prototype.大きさFILTERXML = function (ast, state) {
      var 高 = 1, 幅 = 1;
      try {
        var v = this.filterxml(ast, state);
        if (!(v instanceof CellError)) {
          var d = (v && v.data) || v;
          if (Array.isArray(d)) { 高 = d.length || 1; 幅 = (d[0] && d[0].length) || 1; }
        }
      } catch (e) { /* 出せない時は 1×1 */ }
      return ArraySize ? new ArraySize(幅, 高) : { width: 幅, height: 高 };
    };

    var 何でも = { argumentType: T.ANY };
    Plug.implementedFunctions = {
      'FILTERXML': { method: 'filterxml', sizeOfResultArrayMethod: '大きさFILTERXML',
        parameters: [何でも, 何でも] }
    };
    var 訳 = {};
    for (var k in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k)) 訳[k] = k;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（XML）', e);
      return 0;
    }
    つなぐ.済み = true;
    つなぐ.数 = Object.keys(Plug.implementedFunctions).length;
    return つなぐ.数;
  }
  つなぐ.済み = false;
  つなぐ.数 = 0;
  つなぐ.部品 = null;

  return { つなぐ: つなぐ };
}));
