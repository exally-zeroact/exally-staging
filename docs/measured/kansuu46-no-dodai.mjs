/* kansuu46-no-dodai.mjs — ★★`kansuu46` の 紙と 突き合わせる 時の 土台★★（2026-09-15）
 *
 *  ★★なぜ 1本に するか（★今日 実際に 起きた★）★★
 *    私は `kansuu46` の 紙の 式を 押す 道具を 書く 時、★材料を 自分で 決めました★。
 *      私の 材料 … A1:B5 = 1,3／2,4／5,6／7,8／9,10
 *      ★紙の 材料★ … ★A1:A5 = 1,2,3,4,5★
 *    ⇒ `=IMSUM(A1:A5)` … ★うち 24 ／ 実Excel 15★
 *    ⇒ ★★928本が「合わない」と 出ました★★＝★★偽の 負け★★
 *    ⇒ ★「うちは 半分 負けて いる」と 報せる 一歩 手前★
 *    ★決まりは 前から 在りました★
 *      `feedback_kami_no_zairyou_ga_bunshou_dake_nara_nise_no_make`
 *      ＝★紙の 材料は 紙から 取る／道具が 勝手に 決めない★
 *    ⇒ ★★読んで いたのに 自分で 作りました★★
 *      ＝★★一番 読まれないのは 自分が 書いた 紙★★
 *
 *  ★★式の 置き場も 材料と 同じ 重さ★★
 *    ★式を 材料の 四角の 中に 置くと ★#CYCLE★★（実測・2026-09-15）
 *      `=SUM(A1:A5)`          中に置く→★#CYCLE★ ／ 外に置く→★24★
 *      `=TAKE(A1:A5,2)`       中に置く→★#CYCLE★ ／ 外に置く→★1★
 *      `=CHOOSEROWS(A1:A5,1)` 中に置く→★#CYCLE★ ／ 外に置く→★1★
 *    ⇒ ★気づかなければ ★#CYCLE だらけの 紙★を「今の 答え」として 焼いて いました★
 *    ⇒ ★だから ★置き場（H1）も ここで 決めます★★
 *
 *  ★★出どころ★★
 *    材料・置き場・押し方は ★`kansuu46/awaseru-346.mjs` が 持って いた 物★を
 *    そのまま 切り出しました（★実Excel を 測った 時と 同じ★と そこに 書いて 在る）。
 *
 *  ★★見て いない 事★★
 *    ・★ここは node の 台です★＝★画面の 数では ありません★
 *    ・★答えが 正しいかは 見て いません★＝★押す 道を 揃えるだけ★
 *    ・★`kansuu46` 以外の 紙の 材料は 別★（★その紙を 作った 道具が 持つ★）
 *
 *  使い方:
 *    const 台 = await 建てる();                       // honban-no-michi.mjs
 *    const r = 押す(台, '=SUM(A1:A5)');               // → { 道, 字 }
 */

/* ★★実Excel を 測った 時と 同じ 材料★★（`kansuu46/awaseru-346.mjs` から）
     A1:A5 = 1,2,3,4,5 ／ B1:B5 = 2,4,6,8,10
     D1 = `=DATE(2024,1,1)` ／ D2 = `=DATE(2026,1,1)`
   ★式は H1★（★材料の 外★＝`#CYCLE` に ならない） */
export const 式の列 = 7;          /* H */
export const 式の行 = 0;          /* 1 */

export function 材料() {
  const 表 = [];
  for (let r = 0; r < 6; r++) 表.push([null, null, null, null, null, null, null, null]);
  表[0][0] = 1; 表[1][0] = 2; 表[2][0] = 3; 表[3][0] = 4; 表[4][0] = 5;
  表[0][1] = 2; 表[1][1] = 4; 表[2][1] = 6; 表[3][1] = 8; 表[4][1] = 10;
  表[0][3] = '=DATE(2024,1,1)';
  表[1][3] = '=DATE(2026,1,1)';
  return 表;
}

/* ★誤りの 名前★（エンジンの 言い方 → 実Excel の 言い方） */
const 赤の名 = (t) => ({
  VALUE: '#VALUE!', DIV_BY_ZERO: '#DIV/0!', NUM: '#NUM!', NA: '#N/A',
  NAME: '#NAME?', REF: '#REF!', CYCLE: '#CYCLE!', ERROR: '#ERROR!', SPILL: '#SPILL!',
}[t] || ('#' + t));

/**
 * ★★1つの 式を 本番の 道で 押す★★（JS層 → convertFormula → エンジン）
 * @param {*} 台 `honban-no-michi.mjs` の `建てる()` が 返した 物
 * @param {string} 式
 * @returns {{道:string, 字:string}}
 */
export function 押す(台, 式) {
  const { EF, hf, SID } = 台;
  let 後;
  try { 後 = EF.convertFormula(式); } catch (e) { return { 道: '変換', 字: '★投げた★' }; }

  /* ★★材料を 先に 板へ 入れる★★（★2026-09-15 に 直しました★）
       ★前は JS層を 呼んでから 材料を 入れて いました★
       ⇒★★JS層（`_jsSet` 27個）は ★空の 板★を 読んで いました★★
       ★実測★ `=DSUM(A1:B5,1,A1:A2)` … ★前 0★／★後 2★（実Excel 2）
              9本の D系で ★合うのが 0本 → 5本★に なりました
       ⇒★★全部 「うちが 負けて いる」向きの 偽の 負け★★
       ★本番の 順★ … `book.html` 2873 は ★本を 読み込んだ 後★に `_jsComputeFormula`
       ⇒★だから ★材料が 入って いるのが 本番★ */
  const 表 = 材料();
  表[式の行][式の列] = 後;
  try { hf.setSheetContent(SID, 表); } catch (e) { return { 道: 'エンジン', 字: '★投げた★' }; }

  try {
    const js = EF._jsComputeFormula(0, 式);
    if (js !== null && js !== undefined) return { 道: 'JS層', 字: String(js) };
  } catch (e) { /* JS層が 投げた＝engine へ */ }
  try {
    const v = hf.getCellValue({ sheet: SID, row: 式の行, col: 式の列 });
    if (v && v.type) return { 道: 'エンジン', 字: 赤の名(v.type) };
    if (v === null || v === undefined) return { 道: 'エンジン', 字: '' };
    return { 道: 'エンジン', 字: String(v) };
  } catch (e) { return { 道: 'エンジン', 字: '★投げた★' }; }
}

/* ★★板を 空に する★★＝★溢れた 表を 板が 持ち続ける★ので 押す たびに 積もる
     ★実測（2026-09-15）★
       `=WRAPCOLS(A1:A5,45294)` … ★512MB でも 落ちる★
       `=WRAPROWS(A1:A5,45294)` … ★37 MB★
       （★`=WRAPROWS(A1:A5,3)` は 1 MB★＝★漏れでは なく 45,294幅の 表を 本当に 作る★）
     ★板を 空に すると 34 MB 戻る★（実測）
   ⇒★★上限を 上げずに 済ませる★★（★上限を 上げるのは 合図を 消す 手★） */
export function 板を空に(台) {
  try { 台.hf.setSheetContent(台.SID, [['']]); } catch (e) { /* 気に しない */ }
}
