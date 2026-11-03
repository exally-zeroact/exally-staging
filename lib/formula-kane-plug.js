/* formula-kane-plug.js — ★lib/formula-kane.js を エンジンに 繋ぐ★（2026-09-07）
 *
 *  ★分けてある 理由★＝計算は node で そのまま 試験できる／ここは 繋ぐ 所だけ。
 *  ★形は lib/formula-nokori-plug.js と 同じ★（作る前に 探して 借りた）。
 *
 *  ★★extends で 作る★★（Object.create だと this.evaluateAst が undefined に なる
 *    ＝2026-08-31 に 実測で 踏んだ）
 *
 *  ★★出す 名前は lib/formula-kane.js の 足した名前() が 正本★★
 *    ⇒ ここで 手で 並べ直さない（★2か所に 名簿が 在ると 必ず ずれる★）
 *    ⇒ ODDFPRICE など ★保留の 2個は ここに 書かない★（2026-09-19 ... 4個 → 2個）
 *
 *  見張り: tests/formula-kane.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaKanePlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function つなぐ(H, F) {
    if (!H || !H.FunctionPlugin) return 0;
    if (つなぐ.済み) return つなぐ.数;                 /* ★2回 繋がない★ */
    var CellError = H.CellError, ErrorType = H.ErrorType;
    var T = H.FunctionArgumentType || {};
    if (!CellError || !ErrorType || !T.ANY) return 0;

    function 値(自, ast, state, i) {
      if (!ast.args[i]) return null;
      var v = 自.evaluateAst(ast.args[i], state);
      if (v instanceof CellError) return v;
      if (v === null || v === undefined) return null;
      return v;
    }
    /* ★★借り物は 日付を 包んで 渡します★★（2026-09-18 実測）
       `=DATE(2008,11,11)` を 引数に すると 中では ★`{ val: 39763, format: ... }`★
       ＝★`Number()` は NaN★ ⇒ ★#VALUE!★
       ★数を 直に 渡すと 39763 の まま★＝★同じ 日・同じ 関数・渡し方だけ 違う★
       ⇒★包みを 外してから 数に します★（お金・割合・時刻も 同じ `val` を 持ちます） */
    function 包みを外す(v) {
      if (v !== null && typeof v === 'object' && Object.prototype.hasOwnProperty.call(v, 'val')) return v.val;
      return v;
    }
    function 数(自, ast, state, i) {
      var v = 包みを外す(値(自, ast, state, i));
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      var n = Number(v);
      return isFinite(n) ? n : new CellError(ErrorType.VALUE);
    }
    function 真偽(自, ast, state, i) {
      var v = 包みを外す(値(自, ast, state, i));
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      return String(v).toUpperCase() === 'TRUE';
    }
    function 直す(x) {
      if (x && x.誤り) return new CellError(ErrorType[x.誤り] || ErrorType.NUM);
      if (typeof x === 'number' && !isFinite(x)) return new CellError(ErrorType.NUM);
      return x;
    }
    /* ★引数を 順に 数で 取る★（1つでも 赤なら その まま 返す） */
    function 数たち(自, ast, state, 個) {
      var 出 = [];
      for (var i = 0; i < 個; i++) {
        var v = 数(自, ast, state, i);
        if (v instanceof CellError) return v;
        出.push(v);
      }
      return 出;
    }
    function 包む(個, 呼ぶ) {
      return function (ast, state) {
        var a = 数たち(this, ast, state, 個);
        if (a instanceof CellError) return a;
        try { return 直す(呼ぶ(a)); } catch (e) { return new CellError(ErrorType.NUM); }
      };
    }

    var 日 = function (n) { return F.数から日(n); };

    var Plug = class extends H.FunctionPlugin {
      accrint(ast, state) {
        var a = 数たち(this, ast, state, 7);
        if (a instanceof CellError) return a;
        var 方 = 真偽(this, ast, state, 7);
        if (方 instanceof CellError) return 方;
        try { return 直す(F.経過利息(a[0], a[1], a[2], a[3], a[4], a[5], a[6], 方 === null ? true : 方)); }
        catch (e) { return new CellError(ErrorType.NUM); }
      }
      coupncd(ast, state) {
        var a = 数たち(this, ast, state, 4);
        if (a instanceof CellError) return a;
        try {
          var p = F.次の利払日(日(a[0]), 日(a[1]), a[2]);
          return F.日から数(p.y, p.m, p.d);
        } catch (e) { return new CellError(ErrorType.NUM); }
      }
      couppcd(ast, state) {
        var a = 数たち(this, ast, state, 4);
        if (a instanceof CellError) return a;
        try {
          var p = F.前の利払日(日(a[0]), 日(a[1]), a[2]);
          return F.日から数(p.y, p.m, p.d);
        } catch (e) { return new CellError(ErrorType.NUM); }
      }
      vdb(ast, state) {
        var a = 数たち(this, ast, state, 6);
        if (a instanceof CellError) return a;
        var 切 = 真偽(this, ast, state, 6);
        if (切 instanceof CellError) return 切;
        try { return 直す(F.可変定率(a[0], a[1], a[2], a[3], a[4], a[5] === null ? 2 : a[5], 切 === null ? false : 切)); }
        catch (e) { return new CellError(ErrorType.NUM); }
      }
    };

    /* ★引数を そのまま 渡すだけの 物★（名前と 個数は 1か所で 並べる） */
    var 並べる = [
      ['accrintm', 5, function (a) { return F.満期一括の経過利息(a[0], a[1], a[2], a[3], a[4]); }],
      ['amordegrc', 7, function (a) { return F.仏定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
      ['amorlinc', 7, function (a) { return F.仏定額(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
      ['coupdaybs', 4, function (a) { return F.前からの日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
      ['coupdays', 4, function (a) { return F.期間の日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
      ['coupdaysnc', 4, function (a) { return F.次までの日数(日(a[0]), 日(a[1]), a[2], a[3] === null ? 0 : a[3]); }],
      ['coupnum', 4, function (a) { return F.利払回数(日(a[0]), 日(a[1]), a[2]); }],
      ['disc', 5, function (a) { return F.割引率(a[0], a[1], a[2], a[3], a[4]); }],
      ['duration', 6, function (a) { return F.期間(a[0], a[1], a[2], a[3], a[4], a[5]); }],
      ['intrate', 5, function (a) { return F.利率(a[0], a[1], a[2], a[3], a[4]); }],
      ['mduration', 6, function (a) { return F.修正期間(a[0], a[1], a[2], a[3], a[4], a[5]); }],
      ['price', 7, function (a) { return F.価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
      ['pricedisc', 5, function (a) { return F.割引債の価格(a[0], a[1], a[2], a[3], a[4]); }],
      ['pricemat', 6, function (a) { return F.満期一括の価格(a[0], a[1], a[2], a[3], a[4], a[5]); }],
      ['received', 5, function (a) { return F.受取額(a[0], a[1], a[2], a[3], a[4]); }],
      ['yield', 7, function (a) { return F.利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6]); }],
      ['yielddisc', 5, function (a) { return F.割引債の利回り(a[0], a[1], a[2], a[3], a[4]); }],
      ['yieldmat', 6, function (a) { return F.満期一括の利回り(a[0], a[1], a[2], a[3], a[4], a[5]); }],
      /* ★★2026-09-19 ... ODDL の 2個★★（★保留から 降ろしました★）
           ＝経営者1 の ★275本で 外れ 0本★／わざと 壊したら 258本 赤
           ＝★basis を 省くと 0★（実Excel と 同じ） */
      ['oddlprice', 8, function (a) { return F.最終端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7] === null || a[7] === undefined ? 0 : a[7]); }],
      ['oddlyield', 8, function (a) { return F.最終端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7] === null || a[7] === undefined ? 0 : a[7]); }],
      /* ★★2026-09-19 ... ODDF の 2個★★（★保留から 降ろしました＝保留 0個★）
           ＝経営者1 の ★414本で 外れ 0本★／わざと 壊したら ★53 / 414★
           ＝★basis を 省くと 0★（実Excel と 同じ） */
      ['oddfprice', 9, function (a) { return F.初回端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8] === null || a[8] === undefined ? 0 : a[8]); }],
      ['oddfyield', 9, function (a) { return F.初回端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], a[8] === null || a[8] === undefined ? 0 : a[8]); }],
    ];
    for (var i = 0; i < 並べる.length; i++) {
      Plug.prototype[並べる[i][0]] = 包む(並べる[i][1], 並べる[i][2]);
    }

    var 何でも = { argumentType: T.ANY };
    var 任意 = { argumentType: T.ANY, optionalArg: true };
    var 引数 = function (要る, 任意の数) {
      var 出 = [];
      for (var j = 0; j < 要る; j++) 出.push(何でも);
      for (var k = 0; k < 任意の数; k++) 出.push(任意);
      return 出;
    };

    Plug.implementedFunctions = {
      'ACCRINT':    { method: 'accrint',    parameters: 引数(6, 2) },
      'ACCRINTM':   { method: 'accrintm',   parameters: 引数(4, 1) },
      'AMORDEGRC':  { method: 'amordegrc',  parameters: 引数(6, 1) },
      'AMORLINC':   { method: 'amorlinc',   parameters: 引数(6, 1) },
      'COUPDAYBS':  { method: 'coupdaybs',  parameters: 引数(3, 1) },
      'COUPDAYS':   { method: 'coupdays',   parameters: 引数(3, 1) },
      'COUPDAYSNC': { method: 'coupdaysnc', parameters: 引数(3, 1) },
      'COUPNCD':    { method: 'coupncd',    parameters: 引数(3, 1) },
      'COUPNUM':    { method: 'coupnum',    parameters: 引数(3, 1) },
      'COUPPCD':    { method: 'couppcd',    parameters: 引数(3, 1) },
      'DISC':       { method: 'disc',       parameters: 引数(4, 1) },
      'DURATION':   { method: 'duration',   parameters: 引数(5, 1) },
      'INTRATE':    { method: 'intrate',    parameters: 引数(4, 1) },
      'MDURATION':  { method: 'mduration',  parameters: 引数(5, 1) },
      'PRICE':      { method: 'price',      parameters: 引数(6, 1) },
      'PRICEDISC':  { method: 'pricedisc',  parameters: 引数(4, 1) },
      'PRICEMAT':   { method: 'pricemat',   parameters: 引数(5, 1) },
      'RECEIVED':   { method: 'received',   parameters: 引数(4, 1) },
      'VDB':        { method: 'vdb',        parameters: 引数(5, 2) },
      'YIELD':      { method: 'yield',      parameters: 引数(6, 1) },
      'YIELDDISC':  { method: 'yielddisc',  parameters: 引数(4, 1) },
      'YIELDMAT':   { method: 'yieldmat',   parameters: 引数(5, 1) },
      'ODDLPRICE':  { method: 'oddlprice',  parameters: 引数(7, 1) },
      'ODDFPRICE':  { method: 'oddfprice',  parameters: 引数(8, 1) },
      'ODDFYIELD':  { method: 'oddfyield',  parameters: 引数(8, 1) },
      'ODDLYIELD':  { method: 'oddlyield',  parameters: 引数(7, 1) }
    };

    var 訳 = {};
    for (var k2 in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k2)) 訳[k2] = k2;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（お金）', e);
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
