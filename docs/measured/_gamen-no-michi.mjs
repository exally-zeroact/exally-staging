/* _gamen-no-michi.mjs — ★画面の 事を 測る 時の 決め 4つ★ 2026-09-25
 *
 *  ★★2人で 決めました（Exally1 ＋ 経営者1・2026-09-25）★★
 *    ★『画面の 事を 測る 時は 画面の 関数を 呼ぶ。真似ない』★
 *
 *  ★★なぜ（★1日で 5回 踏みました★）★★
 *    経営者1
 *      ⑴`cell.d` を 読んだ ........... ★描く 所を 真似た★
 *      ⑵`_字の元` を 読んだ .......... ★書式を 掛ける 所を 真似た★
 *      ⑶`activeSheet` に 代入した .... ★板を 開く 所を 真似た★
 *    私
 *      ⑷`####` 450個 ... ★自分で 作った canvas で `cell.d` を 測った★
 *      ⑸`####` 21個 .... ★結合した マスを 1列ぶんの 幅で 測った★
 *    ⇒★5つ とも 同じ 型★＝★お客さんの 道の 「一部だけ」 真似た★
 *    ⇒★どれも 落ちません★＝★もっともらしい 数が 出ます★
 *
 *  ★★決めの 中身（★機械で 守る★）★★
 *    ⑴★測る 道具は 「画面が 呼ぶ 関数」を そのまま 呼ぶ★
 *    ⑵★その 関数が 1つでも 無ければ ★数を 出さずに 止まる★★（★黙って 代わりで 比べない★）
 *    ⑶★出しに 「どの 道で 取ったか」を 全部 書く★（関数の 名を 並べる）
 *    ⑷★「何個に 効いたか」を 数の 隣に 出す★（★0個なら その場で 分かる★）
 *
 *  ★使い方★
 *    import { 道を確かめる, 画面の字を作る } from './_gamen-no-michi.mjs';
 *    await 道を確かめる(page);                      // ★無ければ ここで 止まります★
 *    const 字 = await 画面の字を作る(page, 板名, '行,列');
 */

/** ★画面が 1マスの 字を 出すまでに 通る 関数★（★この 並びが 道です★）
 *  `switchSheet`  ･･･ ★板を 開く★（`activeSheet` の 代入だけでは 台に 流れない）
 *  `_字の元`      ･･･ 書式を 掛ける 前の 値
 *  `_ゼロを隠すか` ･･･ ゼロを 隠す 板か（★`sheets[activeSheet]` を 見ます★）
 *  `_答えは字か`   ･･･ 答えが 字なら 書式を 掛けない
 *  `cW`          ･･･ 列の 幅（★`General` は 幅で 桁が 変わる★）
 *  `_入る字数`    ･･･ その 幅に 何字 入るか
 *  `fmtForDisplay` ･･･ 書式を 掛ける
 */
export const 要る関数 = [
  'switchSheet', '_字の元', '_ゼロを隠すか', '_答えは字か',
  'cW', '_入る字数', 'fmtForDisplay',
];

export const 道の字 = '客の道（switchSheet→_字の元→_ゼロを隠すか→_答えは字か→fmtForDisplay(_入る字数(cW))）';

/** ★要る 関数が 全部 在るか★＝★1つでも 無ければ 数を 出さずに 止まる★ */
export async function 道を確かめる(page) {
  const 無い = await page.evaluate((名たち) => 名たち.filter((n) => typeof window[n] !== 'function'), 要る関数);
  console.log('  ★道★ ' + 道の字);
  if (無い.length) {
    console.log('  ★★画面の 字は 測れません★★＝★要る 関数が ' + 無い.length + '本 在りません★');
    console.log('    ' + 無い.join(' / '));
    console.log('  ⇒★黙って 代わりの 物で 比べません★（2026-09-25 の 決め）');
    process.exit(8);
  }
  console.log('    ⇒★要る 関数 ' + 要る関数.length + '本 とも 在ります★');
  return true;
}

/** ★1マスの「画面に 出る 字」を 画面の 関数で 作る★
 *  返り ･･･ { 字, 隠した, 板を開いた } ／ マスが 無ければ null
 */
export async function 画面の字を作る(page, 板名, 印) {
  return page.evaluate(([な, k]) => {
    const i = (window.sheets || []).findIndex((s) => s.name === な);
    if (i < 0) return null;
    window.switchSheet(i);                 /* ★板を 開く★（代入だけに しない） */
    const sh = window.sheets[i];
    const cell = (sh.data || {})[k];
    if (!cell) return { 字: '', 隠した: false, 板を開いた: true, マス無し: true };
    const raw = window._字の元(cell);
    if (window._ゼロを隠すか(cell, raw)) return { 字: '', 隠した: true, 板を開いた: true };
    if (window._答えは字か(cell)) {
      return { 字: String(raw === undefined || raw === null ? '' : raw), 隠した: false, 板を開いた: true };
    }
    const w = window.cW(+String(k).split(',')[1]);
    return {
      字: String(window.fmtForDisplay(raw, cell.numFmt, window._入る字数(w, raw, cell.numFmt))),
      隠した: false, 板を開いた: true,
    };
  }, [板名, 印]);
}

/** ★「何個に 効いたか」を 数の 隣に 出す★＝★0個なら その場で 分かる★ */
export function 効いた数を出す(名, 効いた, 見た) {
  const 割 = 見た > 0 ? Math.round((効いた / 見た) * 1000) / 10 : 0;
  console.log('  ★' + 名 + ' ... ' + 効いた.toLocaleString() + '個★'
    + '（見た ' + 見た.toLocaleString() + '個 の ' + 割 + '%）'
    + (効いた === 0 ? '  ★★0個＝1つも 効いて いません★★' : ''));
}
