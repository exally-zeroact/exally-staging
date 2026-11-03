/* formula-soto-plug.js — ★外へ 出る 関数を エンジンに 繋ぐ★（2026-09-07）
 *
 *  ★★エンジンの 関数は「待てない」★★
 *    計算エンジン（HyperFormula）の 関数は ★その場で 答えを 返す★決まりです。
 *    でも 外へ 取りに 行くのは ★時間が かかります★。
 *    ⇒★★1回目は「取りに 行っています」を 返し、届いたら もう一度 計算する★★
 *      ①式を 押す → 覚え書きに 無い → ★取りに 行き始める★ → `#N/A` を 返す
 *      ②届く → 覚え書きに しまう → ★再計算を お願いする★
 *      ③もう一度 計算される → 覚え書きに 在る → ★答えが 出る★
 *    ⇒★実Excel は その場で 止まって 待ちます★＝★ここが 実Excel と 違う 所★
 *    ⇒★でも「答えが 出ない」のでは なく「少し 遅れて 出る」★
 *
 *  ★★同じ 住所は 1回しか 取りに 行かない★★
 *    ＝1000行 同じ 式が 在っても 外へ 出るのは 1回（お金と 待ち時間）
 *
 *  ★★外へ 出るのは `api/soto.js` 経由だけ★★
 *    画面から 直に 外へ 出しません（実測で「送るのは 通る」事が 分かっている）
 *
 *  見張り: tests/formula-soto.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaSotoPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★覚え書き★ … 住所（か 頼み文）→ { 状態, 値, 訳 } */
  var 覚え = new Map();
  var 便 = null;          /* 外と やり取りする 道具（画面が 入れる／試験は 作り物を 入れる） */

  function 覚えを消す() { 覚え.clear(); 待ち = 0; }
  function 覚えの数() { return 覚え.size; }

  /* ★★お客さんに「取りに 行っています」を 見せる★★（2026-09-07 指示役）
     ⇒★`#N/A` だけだと お客さんは「★壊れた★」と 思います★
     ⇒ 取りに 行き始めたら 1行 出し、★全部 届いたら 消す★
     ⇒ 数を 持つのは ★何本 待っているか★を 出す為（1本でも 出す） */
  var 待ち = 0;
  function 知らせる(文) {
    if (便 && typeof 便.知らせる === 'function') {
      try { 便.知らせる(文); } catch (e) { /* 知らせられなくても 計算は 続ける */ }
    }
  }

  /* ══ ★★外へ 出して よいか（お客さんが 決めます）★★ ══（2026-09-26）
       ★★なぜ 要るか★★
         この 紙の 頭に こう 書いて あります:
           「もらった `.xlsx` の 式は ★式のまま★ 入る＝★開いた だけで 走る★」
         ⇒★知らない 人の 本を 開いた だけで 外へ 出る★ 事に なります。
         ⇒司さん 09-25「★おすすめで 直せ★」／経営者1 の 推し ⑶
         ⇒★★その 4つが 在る 本だけ お客さんに 訊く★★
       ★★決めの 中身★★
         ・初めは ★null（まだ 訊いて いない）★ ＝ ★出しません★
         ・お客さんが 「はい」と 言った 本だけ 出します
         ・★断られたら 取りに 行きません★＝`#N/A` と 知らせを 出します
         ・★本を 開き直したら また 訊きます★（`外へ出すを忘れる`）
       ★門は 1か所で 足ります★＝★4つ とも `頼む` を 通ります★
         （`webservice` `stockhistory` `translate` `detectlanguage` ･･･ 機械で 数えて 4/4）
       ★見張り★ tests/soto-wo-kiku-webkit.mjs */
  var 外へ出してよい = null;
  function 外へ出す(よいか) { 外へ出してよい = (よいか === true); }
  function 外へ出してよいか() { return 外へ出してよい; }
  function 外へ出すを忘れる() { 外へ出してよい = null; }

  /** ★1回だけ 取りに 行く★（同じ 鍵で 2回 走らせない） */
  function 頼む(鍵, 走る) {
    var 今 = 覚え.get(鍵);
    if (今) return 今;
    /* ★★お客さんが 「はい」と 言って いない 間は 出ません★★（上の 断りを 見て ください）
         ＝★覚えに 入れて おきます★＝同じ 鍵で 何度も 知らせを 出さない 為
         ＝★後で 「はい」に なったら 覚えを 消して 取りに 行けます★ */
    if (外へ出してよい !== true) {
      var 止 = { 状態: '止めた', 値: null, 訳: '外へ 出す 事を まだ 許して いません' };
      覚え.set(鍵, 止);
      知らせる('外へ 出す 関数が 在ります。出して よいか まだ 決まって いません');
      return 止;
    }
    var 札 = { 状態: '取得中', 値: null, 訳: '' };
    覚え.set(鍵, 札);
    待ち++;
    知らせる('外の 物を 取りに 行っています…（' + 待ち + '件）');
    Promise.resolve()
      .then(走る)
      .then(function (v) { 札.状態 = '済'; 札.値 = v; })
      .catch(function (e) { 札.状態 = 'だめ'; 札.訳 = (e && e.message) || '取りに 行けませんでした'; })
      .then(function () {
        待ち--;
        /* ★最後の 1本が 終わった 時だけ 結果を 出す★（1本ごとに 出すと うるさい） */
        if (待ち <= 0) {
          待ち = 0;
          var だめ = 0, 済 = 0;
          覚え.forEach(function (x) { if (x.状態 === 'だめ') だめ++; else if (x.状態 === '済') 済++; });
          知らせる(だめ
            ? '外の 物を ' + だめ + '件 取れませんでした（届いた 分は 出ています）'
            : '外の 物が 届きました（' + 済 + '件）');
        }
        if (便 && typeof 便.再計算 === 'function') { try { 便.再計算(); } catch (e2) { /* 続ける */ } }
      });
    return 札;
  }

  function つなぐ(H, F, 道具) {
    if (!H || !H.FunctionPlugin) return 0;
    便 = 道具 || 便;
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
    function 字(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      return v === null ? '' : String(v);
    }
    function 数(自, ast, state, i) {
      var v = 値(自, ast, state, i);
      if (v instanceof CellError) return v;
      if (v === null || v === '') return null;
      var n = Number(v);
      return isFinite(n) ? n : null;
    }
    /** ★札を 答えに 直す★＝取得中は #N/A・だめは #VALUE! */
    function 札を答えに(札, 直す) {
      /* ★止めた 時も `#N/A`★＝★「取りに 行って いない」と 「取れなかった」を 分けて 覚えます★ */
      if (札.状態 === '止めた') return new CellError(ErrorType.NA);
      if (札.状態 === '取得中') return new CellError(ErrorType.NA);
      if (札.状態 === 'だめ') return new CellError(ErrorType.VALUE);
      var v = 直す ? 直す(札.値) : 札.値;
      if (v && v.誤り) return new CellError(ErrorType[v.誤り] || ErrorType.VALUE);
      return v;
    }
    function 取る(url) {
      if (!便 || typeof 便.取る !== 'function') throw new Error('外へ 出る 道具が 入っていません');
      return 便.取る(url);
    }
    function 聞く(文) {
      if (!便 || typeof 便.聞く !== 'function') throw new Error('AI に 聞く 道具が 入っていません');
      return 便.聞く(文);
    }

    var Plug = class extends H.FunctionPlugin {
      webservice(ast, state) {
        /* ★★住所は「式の 中に 直に 書いた 字」だけ★★（2026-09-07 指示役の 問いで 直した）
           ★なぜ★
             許した 相手だけに しても、`="https://許した相手/?q=" & A1` と 書けたら
             ★A1 の 中身は 相手に 届きます★（行き先は 安全でも ★中身が 出る★）
             ⇒★もらった ファイルに その 式が 仕込まれていたら 同じ事★
           ★直し★
             住所の 所が ★字そのもの（STRING）★で なければ 断る
             ⇒ セルを 指す／つなげる／関数で 作る … 全部 断る
             ⇒★★式を 書いた 時に 決まっている 住所しか 使えない★★
             ⇒★★だから 開いた 人の 中身は 出て行きようが ない★★ */
        var 一 = ast.args && ast.args[0];
        if (!一 || 一.type !== 'STRING') return new CellError(ErrorType.VALUE);
        var u = 字(this, ast, state, 0);
        if (u instanceof CellError) return u;
        if (!F.住所の形か(u)) return new CellError(ErrorType.VALUE);
        var 札 = 頼む('web|' + u, function () { return 取る(u); });
        return 札を答えに(札, function (v) { return String(v == null ? '' : v); });
      }
      stockhistory(ast, state) {
        var s = 字(this, ast, state, 0);
        if (s instanceof CellError) return s;
        var 始 = 数(this, ast, state, 1), 終 = 数(this, ast, state, 2);
        var 何を = 数(this, ast, state, 4);
        var 見出し = 数(this, ast, state, 3);
        var url = F.株の住所(s, 始, 終);
        if (!url) return new CellError(ErrorType.VALUE);
        var 札 = 頼む('stock|' + url, function () { return 取る(url); });
        var 出 = 札を答えに(札, function (v) {
          return F.株の表(v, 何を === null ? 0 : Math.floor(何を), 見出し === null ? 1 : (見出し ? 1 : 0));
        });
        if (出 instanceof CellError) return 出;
        return H.SimpleRangeValue ? H.SimpleRangeValue.onlyValues(出) : 出;
      }
      translate(ast, state) {
        var 文 = 字(this, ast, state, 0);
        if (文 instanceof CellError) return 文;
        var 元 = 字(this, ast, state, 1), 先 = 字(this, ast, state, 2);
        if (元 instanceof CellError) return 元;
        if (先 instanceof CellError) return 先;
        if (文 === '') return '';
        var 頼み = F.訳す頼み(文, 元, 先);
        var 札 = 頼む('ai|' + 頼み, function () { return 聞く(頼み); });
        return 札を答えに(札, function (v) { return F.訳を整える(v); });
      }
      detectlanguage(ast, state) {
        var 文 = 字(this, ast, state, 0);
        if (文 instanceof CellError) return 文;
        if (文 === '') return new CellError(ErrorType.VALUE);
        var 頼み = F.何語か頼み(文);
        var 札 = 頼む('ai|' + 頼み, function () { return 聞く(頼み); });
        return 札を答えに(札, function (v) { return F.言語コードを拾う(v); });
      }
    };

    var ArraySize = H.ArraySize;
    Plug.prototype.大きさSTOCK = function (ast, state) {
      var 高 = 1, 幅 = 2;
      try {
        var v = this.stockhistory(ast, state);
        if (!(v instanceof CellError)) {
          var d = (v && v.data) || v;
          if (Array.isArray(d)) { 高 = d.length || 1; 幅 = (d[0] && d[0].length) || 1; }
        }
      } catch (e) { /* 取れない時は 1×2 */ }
      return ArraySize ? new ArraySize(幅, 高) : { width: 幅, height: 高 };
    };

    var 何でも = { argumentType: T.ANY };
    var 任意 = { argumentType: T.ANY, optionalArg: true };
    /* ══ ★★`isVolatile: true` が 4つ とも 要ります★★ ══（2026-09-26）
         ★★なぜ★★
           この 4つの 答えは ★入れた 物が 変わって いなくても 変わります★。
             ・取りに 行って いる 間 ･･･ `#N/A` ⇒ 届いたら ★本当の 答え★
             ・お客さんが 「つないで よい」を 押した 時 ⇒ ★止めた から 取りに 行くへ★
           ⇒★印が 無いと 台（HyperFormula）が 前の 答えを 覚えた ままに します★
           ⇒★★計算し直しても 答えが 変わりません★★
         ★★実測（2026-09-26・見張りで 出ました）★★
           「つないで よい」を 押した 後 ･･･ 答え ★`#N/A` の まま★ ／ 外へ 出た ★0回★
           ⇒★覚えを 消して 計算し直しても 台が 動きません★
         ★記憶「HyperFormula プラグイン化で 踏む 罠＝INDIRECT は isVolatile 必須」と 同じ 家★
         ⇒★★これは 私の 直しの 前から の 穴です★★
           ＝★届いた 後の 「再計算」も 効いて いなかった 事に なります★
           ＝★但し 司さんの 実物には この 4つが 0個 なので 誰も 気づけませんでした★
         ★見張り★ tests/soto-wo-kiku-webkit.mjs */
    Plug.implementedFunctions = {
      'WEBSERVICE':     { method: 'webservice',     parameters: [何でも], isVolatile: true },
      'STOCKHISTORY':   { method: 'stockhistory',   sizeOfResultArrayMethod: '大きさSTOCK',
        parameters: [何でも, 任意, 任意, 任意, 任意, 任意, 任意], isVolatile: true },
      'TRANSLATE':      { method: 'translate',      parameters: [何でも, 任意, 任意], isVolatile: true },
      'DETECTLANGUAGE': { method: 'detectlanguage', parameters: [何でも], isVolatile: true }
    };

    var 訳 = {};
    for (var k in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k)) 訳[k] = k;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（外）', e);
      return 0;
    }
    つなぐ.済み = true;
    つなぐ.数 = Object.keys(Plug.implementedFunctions).length;
    return つなぐ.数;
  }
  つなぐ.済み = false;
  つなぐ.数 = 0;

  return { つなぐ: つなぐ, 覚えを消す: 覚えを消す, 覚えの数: 覚えの数, 待ちの数: function () { return 待ち; },
    外へ出す: 外へ出す, 外へ出してよいか: 外へ出してよいか, 外へ出すを忘れる: 外へ出すを忘れる,
    道具を入れる: function (d) { 便 = d; } };
}));
