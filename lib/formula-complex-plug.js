/* formula-complex-plug.js — ★IM系の 出す 字だけ 実Excel に 合わせる★（2026-09-08）
 *
 *  ★★計算には 触らない★★
 *    HyperFormula が 持っている ★元の 実装を そのまま 継ぐ★（`getFunctionPlugin`）。
 *    ⇒ 直すのは ★出た 字★だけ＝★壊す 幅が 一番 狭い★
 *
 *  ★★継ぐだけでは 足りない（実物で 踏んだ）★★
 *    HyperFormula は ★自前の prototype★ を 見に 行く。
 *    `class 新 extends 元 {}` で 継いだ だけだと 積み直す 時に こう 落ちる:
 *      ★"Function method complex not found in plugin …"★
 *    ⇒ だから ★26個 全部を 自分の 物として 置き直す★
 *      （21個は 包む ／ 残り 5個は ★元の まま★ 置き直すだけ）
 *
 *  ★★包む 相手を 間違えると 壊れる★★
 *    IM系 25個の うち ★4個は 数（Double）を 返す★ … IMABS IMAGINARY IMARGUMENT IMREAL
 *    実Excel の =IMARGUMENT(-2) は ★3.1415926535897931（17桁）★
 *    ⇒★そこに 15桁を かけると 壊れる★
 *    ⇒★決まり★ 名前が 似ている 物を まとめて 直す 時は ★返す 物の 種類を 先に 数える★
 *
 *  ★★『切る』では ない・『丸める』★★（指示役 2026-09-08）
 *    ⇒ 訳と 実測は `lib/formula-complex.js` の 頭に 書いてある
 *
 *  見張り: tests/formula-complex.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaComplexPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /**
   * ★エンジンに 繋ぐ★（画面から 1回だけ 呼ぶ）
   * @param {*} H HyperFormula の 名前空間
   * @param {*} F lib/formula-complex.js
   * @returns {number} 包んだ 個数（0＝繋げなかった）
   */
  function つなぐ(H, F) {
    if (!H || !F || typeof H.getFunctionPlugin !== 'function') return 0;
    if (typeof H.registerFunctionPlugin !== 'function') return 0;
    if (つなぐ.済み) return つなぐ.数;

    var 元 = H.getFunctionPlugin('IMCOS');
    if (!元 || !元.implementedFunctions || !元.prototype) return 0;

    /* ★元が 持っている 名前を 先に 数える★（★思い込みで 名前を 並べない★） */
    var 名たち = Object.keys(元.implementedFunctions);
    if (!名たち.length) return 0;

    /* ★★名簿に 無い 名前が 1つでも 在ったら 積まない★★
       ⇒★2026-09-08 に 踏んだ★ … 私は「26個」と 書きながら 名簿は 25個しか 無く、
         ★COMPLEX が 名簿に 無いまま 黙って 素通り★していた（指示役が 見つけた）
       ⇒★『そのまま 通す』は ★名簿に 書いてある 時だけ★★
       ⇒ 元に 新しい 関数が 増えた 時も ここで 止まる（＝黙って 素通りしない） */
    var 名簿 = F.文字を返す.concat(F.数を返す, F.まだ測っていない || []);
    var 名簿に無い = [];
    for (var k = 0; k < 名たち.length; k++) {
      if (名簿.indexOf(名たち[k]) < 0) 名簿に無い.push(名たち[k]);
    }
    つなぐ.名簿に無い = 名簿に無い;
    if (名簿に無い.length) return 0;

    /* ★元は ES6 の class★＝`元.apply(this, …)` では 作れない
       （実物で 踏んだ … ★"Class constructor s cannot be invoked without 'new'"★）
       ⇒ `Reflect.construct` で 作る（この ファイルは ES5 の 書き方の まま 保てる） */
    if (typeof Reflect === 'undefined' || typeof Reflect.construct !== 'function') return 0;
    function 十五桁のIM() { return Reflect.construct(元, arguments, 十五桁のIM); }
    十五桁のIM.prototype = Object.create(元.prototype);
    十五桁のIM.prototype.constructor = 十五桁のIM;
    十五桁のIM.implementedFunctions = 元.implementedFunctions;
    /* ★別名（日本語UI 等）も 持っていたら 一緒に 連れて行く★ */
    if (元.aliases) 十五桁のIM.aliases = 元.aliases;

    var 包んだ = 0, そのまま = 0, 見つからない = [];
    for (var i = 0; i < 名たち.length; i++) {
      var 名 = 名たち[i];
      var 手 = 元.implementedFunctions[名].method;
      var 元の手 = 元.prototype[手];
      if (typeof 元の手 !== 'function') { 見つからない.push(名); continue; }
      if (F.文字を返す.indexOf(名) >= 0) {
        /* ★包む★＝元を 呼んでから 出た 字だけ 直す */
        十五桁のIM.prototype[手] = (function (中身) {
          return function () { return F.字を直す(中身.apply(this, arguments)); };
        }(元の手));
        包んだ++;
      } else {
        /* ★そのまま★＝★自前の prototype に 置き直すだけ★（置かないと 積み直しが 落ちる） */
        十五桁のIM.prototype[手] = 元の手;
        そのまま++;
      }
    }
    if (見つからない.length) return 0;          /* ★1つでも 欠けたら 積まない★（黙って 半分に しない） */
    if (包んだ !== F.文字を返す.length) return 0; /* ★21個 全部 包めた 時だけ★ */

    try { H.registerFunctionPlugin(十五桁のIM); } catch (e) { return 0; }

    つなぐ.済み = true;
    つなぐ.数 = 包んだ;
    つなぐ.そのまま = そのまま;
    return 包んだ;
  }

  return { つなぐ: つなぐ };
}));
