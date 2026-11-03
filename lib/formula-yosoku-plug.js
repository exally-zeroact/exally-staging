/* formula-yosoku-plug.js — ★lib/formula-yosoku.js を エンジンに 繋ぐ★（2026-09-07）
 *
 *  ★形は lib/formula-nokori-plug.js と 同じ★（作る前に 探して 借りた）
 *  ★★extends で 作る★★（Object.create だと this.evaluateAst が undefined に なる）
 *
 *  ★★CONVERT は ★JS層から こちらへ 移した★★
 *    ＝exally-formula.js の 中で 4桁に 丸めていた（実測 0.4536／実Excel 0.45359237）
 *    ⇒★1つの 関数は 1か所でだけ 定義する★（両方に 居ると 先に 当たった 方が 勝つ）
 *
 *  見張り: tests/formula-yosoku.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaYosokuPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function つなぐ(H, F, 場) {
    if (場) つなぐ.場 = 場;
    if (!H || !H.FunctionPlugin) return 0;
    if (つなぐ.済み) return つなぐ.数;
    var CellError = H.CellError, ErrorType = H.ErrorType;
    var T = H.FunctionArgumentType || {};
    if (!CellError || !ErrorType || !T.ANY) return 0;

    function 表に(v) {
      if (v && typeof v.data !== 'undefined' && Array.isArray(v.data)) return v.data;
      if (v && typeof v.simpleRangeValue !== 'undefined') return 表に(v.simpleRangeValue);
      if (v && typeof v.raw === 'function') { try { return v.raw(); } catch (e) { /* 続ける */ } }
      return F.表にする(v);
    }
    function 値(自, ast, state, i) {
      if (!ast.args[i]) return null;
      var v = 自.evaluateAst(ast.args[i], state);
      if (v instanceof CellError) return v;
      if (v === null || v === undefined) return null;
      return v;
    }
    function 表(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      if (v === null) return null;
      return 表に(v);
    }
    function 数(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      return v;
    }
    function 真偽(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      return String(v).toUpperCase() === 'TRUE';
    }
    function 直す(x) {
      if (x && x.誤り) return new CellError(ErrorType[x.誤り] || ErrorType.VALUE);
      if (typeof x === 'number' && !isFinite(x)) return new CellError(ErrorType.NUM);
      return x;
    }
    function 範囲に(x) {
      if (x && x.誤り) return new CellError(ErrorType[x.誤り] || ErrorType.VALUE);
      /* ★★表の 中の マスの 誤りも 直す（2026-09-09 に 足した）★★
         ★前は 外側の 誤りしか 見て いませんでした★
         ⇒ 表の 中に `{誤り:'NA'}` が 在ると ★そのまま 素通り★し、
           画面に ★`[object Object]`★ と 出て いた（LINEST の 埋め物で 踏んだ）
         ⇒★誤りは 表の 中でも 誤りに する★ */
      if (Array.isArray(x)) {
        x = x.map(function (行) {
          return Array.isArray(行) ? 行.map(function (v) {
            return (v && v.誤り) ? new CellError(ErrorType[v.誤り] || ErrorType.VALUE) : v;
          }) : 行;
        });
      }
      return H.SimpleRangeValue ? H.SimpleRangeValue.onlyValues(x) : x;
    }

    var Plug = class extends H.FunctionPlugin {
      convert(ast, state) {
        var n = 数(this, ast, state, 0);
        if (n instanceof CellError) return n;
        var a = 数(this, ast, state, 1), b = 数(this, ast, state, 2);
        if (a instanceof CellError) return a;
        if (b instanceof CellError) return b;
        return 直す(F.単位を変える(n, a === null ? '' : a, b === null ? '' : b));
      }
      erfprecise(ast, state) {
        var x = 数(this, ast, state, 0);
        if (x instanceof CellError) return x;
        return 直す(F.誤差(x === null ? 0 : Number(x)));
      }
      erfcprecise(ast, state) {
        var x = 数(this, ast, state, 0);
        if (x instanceof CellError) return x;
        return 直す(F.誤差の残り(x === null ? 0 : Number(x)));
      }
      percentof(ast, state) {
        var a = 表(this, ast, state, 0), b = 表(this, ast, state, 1);
        if (a instanceof CellError) return a;
        if (b instanceof CellError) return b;
        return 直す(F.割合(a, b));
      }
      trend(ast, state) {
        var y = 表(this, ast, state, 0);
        if (y instanceof CellError) return y;
        var x = 表(this, ast, state, 1), nx = 表(this, ast, state, 2);
        if (x instanceof CellError) return x;
        if (nx instanceof CellError) return nx;
        var c = 真偽(this, ast, state, 3);
        if (c instanceof CellError) return c;
        return 範囲に(F.直線予測(y, x, nx, c));
      }
      growth(ast, state) {
        var y = 表(this, ast, state, 0);
        if (y instanceof CellError) return y;
        var x = 表(this, ast, state, 1), nx = 表(this, ast, state, 2);
        if (x instanceof CellError) return x;
        if (nx instanceof CellError) return nx;
        var c = 真偽(this, ast, state, 3);
        if (c instanceof CellError) return c;
        return 範囲に(F.増える予測(y, x, nx, c));
      }
      /* ★LINEST（2026-09-09 に 足した）★
         今までは exally-formula.js の JS層が ★裸の =LINEST(範囲,範囲) だけ★を 拾い、
         ★傾き 1つ★しか 返して いなかった（入れ子は #NAME? ／ x 2本は 静かに 違う 答え）。
         ⇒★LOGEST と 同じ 土台で 表として 返す★ */
      linest(ast, state) {
        var y = 表(this, ast, state, 0);
        if (y instanceof CellError) return y;
        var x = 表(this, ast, state, 1);
        if (x instanceof CellError) return x;
        var c = 真偽(this, ast, state, 2), st = 真偽(this, ast, state, 3);
        if (c instanceof CellError) return c;
        if (st instanceof CellError) return st;
        return 範囲に(F.直線の係数(y, x, c, st));
      }
      logest(ast, state) {
        var y = 表(this, ast, state, 0);
        if (y instanceof CellError) return y;
        var x = 表(this, ast, state, 1);
        if (x instanceof CellError) return x;
        var c = 真偽(this, ast, state, 2), s = 真偽(this, ast, state, 3);
        if (c instanceof CellError) return c;
        if (s instanceof CellError) return s;
        return 範囲に(F.増えるの係数(y, x, c, s));
      }
      info(ast, state) {
        var k = 数(this, ast, state, 0);
        if (k instanceof CellError) return k;
        /* ★場の 事は 外から 入れる★（画面が 入れる／試験は 作り物）
           ⇒ 入っていなければ ★分かる 範囲だけ★で 答える（嘘を 作らない） */
        var 場 = (つなぐ.場 && typeof つなぐ.場 === 'function') ? つなぐ.場() : (つなぐ.場 || {});
        return 直す(F.場の事(k === null ? '' : k, 場));
      }
      randarray(ast, state) {
        var r = 数(this, ast, state, 0), c = 数(this, ast, state, 1);
        var lo = 数(this, ast, state, 2), hi = 数(this, ast, state, 3);
        var w = 真偽(this, ast, state, 4);
        for (var i = 0, a = [r, c, lo, hi, w]; i < a.length; i++) {
          if (a[i] instanceof CellError) return a[i];
        }
        return 範囲に(F.でたらめの表(r, c, lo, hi, w));
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
    Plug.prototype.大きさTREND = 大きさ('trend');
    Plug.prototype.大きさGROWTH = 大きさ('growth');
    Plug.prototype.大きさLINEST = 大きさ('linest');
    Plug.prototype.大きさLOGEST = 大きさ('logest');
    Plug.prototype.大きさRANDARRAY = 大きさ('randarray');

    var 何でも = { argumentType: T.ANY };
    var 任意 = { argumentType: T.ANY, optionalArg: true };
    Plug.implementedFunctions = {
      'CONVERT':      { method: 'convert',      parameters: [何でも, 何でも, 何でも] },
      'ERF.PRECISE':  { method: 'erfprecise',   parameters: [何でも] },
      'ERFC.PRECISE': { method: 'erfcprecise',  parameters: [何でも] },
      'PERCENTOF':    { method: 'percentof',    parameters: [何でも, 何でも] },
      'INFO':         { method: 'info',         parameters: [何でも] },
      'TREND':        { method: 'trend',        sizeOfResultArrayMethod: '大きさTREND',
        parameters: [何でも, 任意, 任意, 任意] },
      'GROWTH':       { method: 'growth',       sizeOfResultArrayMethod: '大きさGROWTH',
        parameters: [何でも, 任意, 任意, 任意] },
      'LINEST':       { method: 'linest',       sizeOfResultArrayMethod: '大きさLINEST',
        parameters: [何でも, 任意, 任意, 任意] },
      'LOGEST':       { method: 'logest',       sizeOfResultArrayMethod: '大きさLOGEST',
        parameters: [何でも, 任意, 任意, 任意] },
      'RANDARRAY':    { method: 'randarray',    sizeOfResultArrayMethod: '大きさRANDARRAY',
        parameters: [任意, 任意, 任意, 任意, 任意] }
    };

    var 訳 = {};
    for (var k2 in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k2)) 訳[k2] = k2;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（予測）', e);
      return 0;
    }
    つなぐ.済み = true;
    つなぐ.数 = Object.keys(Plug.implementedFunctions).length;
    return つなぐ.数;
  }
  つなぐ.済み = false;
  つなぐ.数 = 0;
  つなぐ.場 = {};

  return { つなぐ: つなぐ };
}));
