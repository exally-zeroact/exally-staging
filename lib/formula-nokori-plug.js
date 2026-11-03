/* formula-nokori-plug.js — ★lib/formula-nokori.js を エンジンに 繋ぐ★（2026-09-06）
 *
 *  ★分けてある 理由★＝計算は node で そのまま 試験できる／ここは 繋ぐ 所だけ。
 *  ★形は lib/formula-extra-plug.js と 同じ★（作る前に 探して 借りた）。
 *
 *  ★★extends で 作る★★（Object.create だと this.evaluateAst が undefined に なる
 *    ＝2026-08-31 に 実測で 踏んだ）
 *
 *  見張り: tests/formula-nokori-plug.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaNokoriPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function つなぐ(H, F) {
    if (!H || !H.FunctionPlugin) return 0;
    if (つなぐ.済み) return つなぐ.数;                 /* ★2回 繋がない★ */
    var CellError = H.CellError, ErrorType = H.ErrorType;
    var T = H.FunctionArgumentType || {};
    if (!CellError || !ErrorType || !T.ANY) return 0;

    function 表に(v) {
      if (v && typeof v.data !== 'undefined' && Array.isArray(v.data)) return v.data;
      if (v && typeof v.simpleRangeValue !== 'undefined') return 表に(v.simpleRangeValue);
      if (v && typeof v.raw === 'function') { try { return v.raw(); } catch (e) { /* 続ける */ } }
      return F.表にする(v);
    }
    function 直す(x) {
      if (x && x.誤り) return new CellError(ErrorType[x.誤り] || ErrorType.VALUE);
      return x;
    }
    function 範囲に(x) {
      if (x && x.誤り) return new CellError(ErrorType[x.誤り] || ErrorType.VALUE);
      return H.SimpleRangeValue ? H.SimpleRangeValue.onlyValues(x) : x;
    }
    function 値(自, ast, state, i) {
      if (!ast.args[i]) return null;
      var v = 自.evaluateAst(ast.args[i], state);
      if (v instanceof CellError) return v;
      if (v === null || v === undefined) return null;
      if (v && typeof v.data !== 'undefined') return v;
      return v;
    }
    function 字(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      return v === null ? '' : String(v);
    }
    function 数(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      return Number(v);
    }

    var Plug = class extends H.FunctionPlugin {
      findb(ast, state) {
        var a = 字(this, ast, state, 0), b = 字(this, ast, state, 1), c = 数(this, ast, state, 2);
        if (a instanceof CellError) return a;
        if (b instanceof CellError) return b;
        if (c instanceof CellError) return c;
        return 直す(F.探すB(a, b, c === null ? 1 : c));
      }
      searchb(ast, state) {
        var a = 字(this, ast, state, 0), b = 字(this, ast, state, 1), c = 数(this, ast, state, 2);
        if (a instanceof CellError) return a;
        if (b instanceof CellError) return b;
        if (c instanceof CellError) return c;
        return 直す(F.探すB大小なし(a, b, c === null ? 1 : c));
      }
      replaceb(ast, state) {
        var s = 字(this, ast, state, 0), i = 数(this, ast, state, 1);
        var n = 数(this, ast, state, 2), t = 字(this, ast, state, 3);
        for (var k = 0; k < 4; k++) { /* 誤りは そのまま 返す */ }
        if (s instanceof CellError) return s;
        if (i instanceof CellError) return i;
        if (n instanceof CellError) return n;
        if (t instanceof CellError) return t;
        return 直す(F.入れ替えB(s, i, n, t));
      }
      textsplit(ast, state) {
        var s = 字(this, ast, state, 0), 横 = 字(this, ast, state, 1), 縦 = 字(this, ast, state, 2);
        if (s instanceof CellError) return s;
        if (横 instanceof CellError) return 横;
        if (縦 instanceof CellError) return 縦;
        return 範囲に(F.区切って分ける(s, 横, 縦 === '' ? null : 縦));
      }
      regextest(ast, state) {
        var s = 字(this, ast, state, 0), p = 字(this, ast, state, 1), m = 数(this, ast, state, 2);
        if (s instanceof CellError) return s;
        if (p instanceof CellError) return p;
        return 直す(F.正規で調べる(s, p, m === 1));
      }
      regexextract(ast, state) {
        var s = 字(this, ast, state, 0), p = 字(this, ast, state, 1), m = 数(this, ast, state, 3);
        if (s instanceof CellError) return s;
        if (p instanceof CellError) return p;
        return 直す(F.正規で取り出す(s, p, m === 1));
      }
      regexreplace(ast, state) {
        var s = 字(this, ast, state, 0), p = 字(this, ast, state, 1), r = 字(this, ast, state, 2);
        var m = 数(this, ast, state, 4);
        if (s instanceof CellError) return s;
        if (p instanceof CellError) return p;
        if (r instanceof CellError) return r;
        return 直す(F.正規で入れ替える(s, p, r, m === 1));
      }
      sortby(ast, state) {
        var 表 = 表に(this.evaluateAst(ast.args[0], state));
        var 鍵 = 表に(this.evaluateAst(ast.args[1], state));
        var 向 = 数(this, ast, state, 2);
        if (向 instanceof CellError) return 向;
        return 範囲に(F.別の列で並べる(表, 鍵, 向 === null ? 1 : 向));
      }
      munit(ast, state) {
        var n = 数(this, ast, state, 0);
        if (n instanceof CellError) return n;
        return 範囲に(F.単位行列(n));
      }
      minverse(ast, state) {
        var 表 = 表に(this.evaluateAst(ast.args[0], state));
        return 範囲に(F.逆行列(表));
      }
      percentrankinc(ast, state) { return this._順位(ast, state, true); }
      percentrankexc(ast, state) { return this._順位(ast, state, false); }
      _順位(ast, state, 含むか) {
        var 表 = 表に(this.evaluateAst(ast.args[0], state));
        var x = 数(this, ast, state, 1), k = 数(this, ast, state, 2);
        if (x instanceof CellError) return x;
        if (k instanceof CellError) return k;
        return 直す(F.順位の割合(表, x, k === null ? undefined : k, 含むか));
      }
      prob(ast, state) {
        var v = 表に(this.evaluateAst(ast.args[0], state));
        var p = 表に(this.evaluateAst(ast.args[1], state));
        var l = 数(this, ast, state, 2), u = 数(this, ast, state, 3);
        if (l instanceof CellError) return l;
        if (u instanceof CellError) return u;
        return 直す(F.確率(v, p, l, u === null ? undefined : u));
      }
      errortype(ast, state) {
        var v = this.evaluateAst(ast.args[0], state);
        if (v instanceof CellError) {
          var 名 = String(v.type || '').toUpperCase();
          return 直す(F.誤りを番号に(名));
        }
        return new CellError(ErrorType.NA);
      }
    };

    /* ★★大きさは ★ArraySize の 形★で 返す★★（ただの {height,width} では 通らない）
       実測（2026-09-07）… ただの 物を 返していた 間、
         `=MUNIT(3)` を そのまま セルに 入れると ★#ERROR!「Invalid range size」★
         `=INDEX(MUNIT(3),1,1)` は 動くので ★試験では 見つからなかった★
       ⇒★★『中に 入れると 動く／そのまま 入れると 死ぬ』＝一番 見つけにくい★★ */
    var ArraySize = H.ArraySize;
    function 大きさ(名) {
      return function (ast, state) {
        var 高 = 1, 幅 = 1;
        try {
          var v = this[名](ast, state);
          if (!(v instanceof CellError)) {
            var d = (v && v.data) || F.表にする(v);
            高 = d.length || 1;
            幅 = (d[0] && d[0].length) || 1;
          }
        } catch (e) { /* 大きさが 出せない時は 1×1 */ }
        return ArraySize ? new ArraySize(幅, 高) : { width: 幅, height: 高 };
      };
    }
    Plug.prototype.大きさTEXTSPLIT = 大きさ('textsplit');
    Plug.prototype.大きさSORTBY = 大きさ('sortby');
    Plug.prototype.大きさMUNIT = 大きさ('munit');
    Plug.prototype.大きさMINVERSE = 大きさ('minverse');

    var 何でも = { argumentType: T.ANY };
    var 任意 = { argumentType: T.ANY, optionalArg: true };
    Plug.implementedFunctions = {
      'FINDB':           { method: 'findb',    parameters: [何でも, 何でも, 任意] },
      'SEARCHB':         { method: 'searchb',  parameters: [何でも, 何でも, 任意] },
      'REPLACEB':        { method: 'replaceb', parameters: [何でも, 何でも, 何でも, 何でも] },
      'TEXTSPLIT':       { method: 'textsplit', sizeOfResultArrayMethod: '大きさTEXTSPLIT',
        parameters: [何でも, 任意, 任意, 任意, 任意, 任意] },
      'REGEXTEST':       { method: 'regextest',    parameters: [何でも, 何でも, 任意] },
      'REGEXEXTRACT':    { method: 'regexextract', parameters: [何でも, 何でも, 任意, 任意] },
      'REGEXREPLACE':    { method: 'regexreplace', parameters: [何でも, 何でも, 何でも, 任意, 任意] },
      'SORTBY':          { method: 'sortby',   sizeOfResultArrayMethod: '大きさSORTBY',
        parameters: [何でも, 何でも, 任意], repeatLastArgs: 2 },
      'MUNIT':           { method: 'munit',    sizeOfResultArrayMethod: '大きさMUNIT',    parameters: [何でも] },
      'MINVERSE':        { method: 'minverse', sizeOfResultArrayMethod: '大きさMINVERSE', parameters: [何でも] },
      'PERCENTRANK.INC': { method: 'percentrankinc', parameters: [何でも, 何でも, 任意] },
      'PERCENTRANK.EXC': { method: 'percentrankexc', parameters: [何でも, 何でも, 任意] },
      'PROB':            { method: 'prob',      parameters: [何でも, 何でも, 何でも, 任意] },
      'ERROR.TYPE':      { method: 'errortype', parameters: [何でも] }
    };

    var 訳 = {};
    for (var k2 in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k2)) 訳[k2] = k2;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（残り）', e);
      return 0;
    }
    つなぐ.済み = true;
    つなぐ.数 = Object.keys(Plug.implementedFunctions).length;
    return つなぐ.数;
  }
  つなぐ.済み = false;
  つなぐ.数 = 0;

  return { つなぐ: つなぐ };
}));
