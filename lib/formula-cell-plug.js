/* formula-cell-plug.js — ★CELL を エンジンに 繋ぐ★（2026-09-07）
 *
 *  ★★CELL は 他の 関数と 作りが 違う★★
 *    ふつうの 関数は「★中身★」を もらえば 足ります。
 *    CELL は ★どの マスか★（番地・列・行）と ★どう 見えているか★
 *    （幅・表示形式・そろえ）を 聞かれます。
 *    ⇒★中身を 受け取る 作りでは 答えられない★
 *      （`this.evaluateAst(...)` は 中身に 変えてしまう）
 *    ⇒ だから ★式の 形（ast）から 指している 番地を 取り出す★
 *      `ast.args[1].reference.toSimpleCellAddress(state.formulaAddress)`
 *
 *  ★★見た目は 画面から 入れる★★（INFO と 同じ 作り）
 *    エンジンは ★幅も 表示形式も 知りません★（中身しか 持っていない）。
 *    ⇒ `つなぐ(H, F, 見た目)` の `見た目(シートの番号, 行, 列)` を
 *      ★画面が 入れる★／試験は 作り物を 入れる。
 *    ⇒★入っていない 時は 既定の 見た目★（＝実Excel の まっさらな シートと 同じ 答え）
 *
 *  見張り: tests/formula-cell.test.mjs
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.FormulaCellPlug = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function つなぐ(H, F, 見た目) {
    if (!H || !H.FunctionPlugin) return 0;
    if (見た目) つなぐ.見た目 = 見た目;
    if (つなぐ.済み) return つなぐ.数;
    var CellError = H.CellError, ErrorType = H.ErrorType;
    var T = H.FunctionArgumentType || {};
    if (!CellError || !ErrorType || !T.ANY) return 0;

    /** ★指している 番地を 取り出す★（中身に 変えない） */
    function 番地を取る(ast, state) {
      var a = ast.args[1];
      if (!a) return { 場: state.formulaAddress, 自分: true };   /* ★2つ目が 無い＝今 いる マス★（実測） */
      if (a.type === 'CELL_REFERENCE' && a.reference && a.reference.toSimpleCellAddress) {
        return { 場: a.reference.toSimpleCellAddress(state.formulaAddress), 自分: false };
      }
      /* ★範囲を 渡すと 左上★（実測 =CELL("address",A1:C3) → $A$1） */
      if (a.type === 'CELL_RANGE' && a.start && a.start.toSimpleCellAddress) {
        return { 場: a.start.toSimpleCellAddress(state.formulaAddress), 自分: false };
      }
      return null;                                              /* マスを 指していない */
    }

    /* ★既定の 見た目は 中身（lib/formula-cell.js）が 正本★
       ＝2026-09-14 に ここから 移しました。★2か所に 置くと 片方だけ 直します★
         （実際に 自前の 皮が 持たず、CELL 9本が #VALUE! に なって いました） */
    var 既定の見た目 = F.既定の見た目;
    function 見た目を聞く(場) {
      var f = つなぐ.見た目;
      var v = null;
      if (typeof f === 'function') {
        try { v = f(場.sheet, 場.row, 場.col); } catch (e) { v = null; }
      } else if (f && typeof f === 'object') v = f;
      var 出 = {};
      for (var k in 既定の見た目) {
        if (Object.prototype.hasOwnProperty.call(既定の見た目, k)) {
          出[k] = (v && v[k] !== undefined && v[k] !== null) ? v[k] : 既定の見た目[k];
        }
      }
      return 出;
    }

    var Plug = class extends H.FunctionPlugin {
      cell(ast, state) {
        /* ①何を 聞かれたか（字だけ・大文字小文字は 問わない＝実測） */
        var 何を = this.evaluateAst(ast.args[0], state);
        if (何を instanceof CellError) return 何を;
        if (何を === null || 何を === undefined) return new CellError(ErrorType.VALUE);
        if (typeof 何を !== 'string') return new CellError(ErrorType.VALUE);

        /* ②どの マスか */
        var 指 = 番地を取る(ast, state);
        if (!指) return new CellError(ErrorType.VALUE);
        var 場 = 指.場;

        /* ③中身（★空か どうかを 見分ける★＝空は Symbol で 返る） */
        var 中身 = null, 空か = true;
        try {
          var v = this.dependencyGraph.getCellValue(場);
          if (typeof v === 'symbol' || v === undefined) { 中身 = null; 空か = true; }
          else { 中身 = v; 空か = false; }
        } catch (e) { 中身 = null; 空か = true; }
        /* ★誤りの マスも 中身★（実測 =CELL("contents",<=1/0 の マス>) は #DIV/0!） */

        /* ④どう 見えているか（画面が 知っている） */
        var 見 = 見た目を聞く(場);

        var 答 = F.セルの事(何を, {
          行: 場.row, 列: 場.col,
          別のシートか: !指.自分 && 場.sheet !== state.formulaAddress.sheet,
          シート名: 見.シート名, ブック名: 見.ブック名,
          中身: (中身 instanceof CellError) ? null : 中身,
          空か: 空か,
          表示形式: 見.表示形式, そろえ: 見.そろえ,
          幅の点: 見.幅の点, 既定の点: 見.既定の点, 隠れ: 見.隠れ,
          ロックなし: !!見.ロックなし
        });
        if (答 && 答.誤り) return new CellError(ErrorType[答.誤り] || ErrorType.VALUE);
        /* ★中身を 聞かれて そのマスが 誤りなら その 誤りを そのまま 返す★ */
        if (中身 instanceof CellError) {
          var k = String(何を).trim().toLowerCase();
          if (k === 'contents') return 中身;
        }
        return 答;
      }
    };

    var 何でも = { argumentType: T.ANY };
    /* ★2つ目は 無くてもよい★（実測 =CELL("row") は 今 いる マスの 行） */
    var 無くてもよい = { argumentType: T.ANY, optionalArg: true };
    Plug.implementedFunctions = {
      'CELL': { method: 'cell', parameters: [何でも, 無くてもよい]  }
    };
    var 訳 = {};
    for (var k in Plug.implementedFunctions) {
      if (Object.prototype.hasOwnProperty.call(Plug.implementedFunctions, k)) 訳[k] = k;
    }
    try {
      H.registerFunctionPlugin(Plug, { enGB: 訳 });
    } catch (e) {
      if (typeof console !== 'undefined') console.warn('関数を 足せませんでした（CELL）', e);
      return 0;
    }
    つなぐ.済み = true;
    つなぐ.数 = Object.keys(Plug.implementedFunctions).length;
    return つなぐ.数;
  }
  つなぐ.済み = false;
  つなぐ.数 = 0;
  つなぐ.見た目 = null;

  return { つなぐ: つなぐ };
}));
