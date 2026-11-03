/* unused-param.test.mjs — ★「持っているのに 渡していない」が 増えていないか★ 2026-08-30
 *
 *  ★なぜ★（監査役の 宿題・2026-08-30）
 *    `startEdit(r, c, init)` は 打った 文字を 受け取る 口を 持っているのに
 *    ★呼んでいる 所が 1つも 渡していなかった★＝★「持っているのに 渡していない」★
 *    ⇒ ★1つ 見つかった 型は だいたい 他にも 在る★ ので 機械で 数える。
 *
 *  ★数え方は 道具 1本★ … `tools/unused-param.mjs` の `数える()` を そのまま 使う
 *    ★試験が 自前で 数えると 道具と ずれる★（二重管理）ので ★読み込んで 呼ぶ★。
 *
 *  ★2026-08-30 に 1件ずつ 中身を 見た 記録★
 *    ・候補 13個 → `startEdit` の ★死んだ 口を 消した★ … 12個
 *    ・道具を 直した（★注記外しを 共通部品に★／★名前で 呼ぶ 形も 見る★） … ★11個★
 *    ・11個は ★全部 正常★（任意の 口／道具の 数え過ぎ）＝docs/EXCEL_PARITY.md の 表
 *
 *  ★道具の 直しで 2回 外した（記録）★
 *    ① 自前の 注記外しが ★HTML の onclick を 行ごと 消した★
 *       ⇒ `applyBorderAll(1)` を 見落として ★生きている 口を 死んでいると 誤報★
 *    ② `呼ぶ('名', 引数)` の 形を 見ておらず ★ピボットの窓を開く を 一度 消しかけた★
 *
 *  走らせ方: node tests/unused-param.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const NL = String.fromCharCode(10);
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

console.log('★「持っているのに 渡していない」★');

/* ★道具を 読み込んで 呼ぶ★（自前で 数えない＝ずれない） */
const 道具の道 = path.join(ROOT, 'tools/unused-param.mjs');
言う(fs.existsSync(道具の道), '★道具が 在る（tools/unused-param.mjs）★');
const 道具 = await import(pathToFileURL(道具の道).href);
言う(typeof 道具.数える === 'function', '★道具が 数を 出せる形に なっている★');

const { 全体, 見つけ } = 道具.数える();
const 一覧 = 見つけ.map((v) => v.ファイル + ':' + v.行 + '  ' + v.名
  + '(' + v.口の名 + ')  口' + v.口 + '/最大' + v.最大);

言う(全体 >= 500, '★空振りしていない★（引数を 持つ 関数 ' + 全体 + '個 を 見た）',
  '500個 未満＝ソースを 読めていない');

/* ── ① 増えていないか ── */
/* ★上限★＝1件ずつ 中身を 見て ★全部 正常★ と 確かめた 数
     08-30 … 11個
     08-31 … ★→ 16個★（5個 増えた）
       増えた 5個：今日 作った 部品の ★数える()★
         ・lib/ctx-menu.js           数える(働き, 窓)
         ・lib/ribbon-keytips.js     数える(根)
         ・lib/ribbon.js             数える(spec)
         ・lib/ribbon-launcher.js    数える(働き)
         ・lib/ribbon-context-spec.js 数える(働き)
       ★どれも 試験から 引数つきで 呼んでいる★（実際に 見た）。
       道具は book.html と lib/*.js しか 見ないので
       ★試験からの 呼び出しが 見えていないだけ★（道具の 数え過ぎ）。
     09-04 … ★→ 15個★（1個 減った）
       診断2本目（ほかの表の その行）で 画面を 直した時に
       ★使っていなかった 口が 1つ 実際に 使われるように なった★。
       ★減ったら 上限も 下げる（戻せない）★＝これが 積み上げの 決まり。
       ★測り方★ … 直す前の 木で 走らせて 16／直した後で 15（両方 実測） */
/* ★★2026-09-18 ... ★見る 範囲が 変わりました（上限は そのまま 15）★★★
     ★前の 範囲★ ... `book.html` `hub.html` ＋ `lib/*.js`
     ★今の 範囲★ ... ★画面の `script src` から 取る★（`exally-formula.js` `js/*.js` が 入った）
     ★なぜ★ ... `lib/named-ranges.js:85 開く(式, シート)` を
                「どこからも 呼ばれない」と ★誤報★して いました
                実物は `exally-formula.js:1221` の `window.名前の箱.開く(f, シート)`
                ＝★repo の 根に 在り 範囲の 外★でした（経営者1 が 実物を 引いて 見つけた）
     ★★範囲を 広げた ので 中身が 入れ替わりました★★（★上限は 1も 上げて いません★）
       ★出た★ ... `lib/named-ranges.js 開く`（★誤報だった★）
                  `lib/shiki-ita-awase.js 数える`（★私が `数える` と いう
                    ★repo に 5つ 在る 名前★を 使った ので 門が 狂った／`写しを数える` に 改名★）
       ★入った★ ... ★どちらも 新しく 見えた ★本物★です★
         `exally-formula.js:1212 convertFormula(f, シート)` 口2/最大1
            ⇒★`シート` を 誰も 渡して いません★＝名前の 開き方が シートを 見て いない
         `js/recipe-store.js:87 覚えた物(opt)` 口1/最大0
       ⇒★★この 2つは まだ 直して いません★★（★別の 1押し★・★隠して いません★） */
const 上限 = 15;
言う(見つけ.length <= 上限,
  '★候補が ' + 上限 + '個から 増えていない（今 ' + 見つけ.length + '個）★',
  '★増えた＝1件ずつ 中身を 見て docs/EXCEL_PARITY.md の 表を 直す★' + NL
    + '       ' + 一覧.join(NL + '       '));

/* ── ② 直した 物は もう 候補に 入らない ── */
for (const 名 of ['startEdit', 'ピボットの窓を開く', 'applyBorderAll']) {
  const 当 = 見つけ.filter((v) => v.名 === 名);
  言う(当.length === 0, '★' + 名 + ' は 候補に 入っていない★',
    当.map((v) => v.ファイル + ':' + v.行).join(' / '));
}

/* ── ③ 道具が 自前の 注記外しを 持っていない ── */
const 道具の字 = fs.readFileSync(道具の道, 'utf8');
言う(/scripts\/lib\/chuki\.mjs/.test(道具の字),
  '★道具は 共通の 注記外しを 使っている（自前で 書いていない）★',
  '★自前だと HTML の onclick を 行ごと 消す（2026-08-30 実際に 踏んだ）★');
言う(/名前で呼ぶ/.test(道具の字),
  '★道具は「名前で 呼ぶ」形も 見ている★',
  '★見ないと 生きている 口を 死んでいると 誤報する★');

/* ── ④ 直した 印が 残っているか ── */
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
言う(/function startEdit\(r,c\)\{/.test(book),
  '★startEdit の 口は 2つ（死んだ 口を 消した）★',
  '★init を 戻すなら 渡す 所も 作る事★');
言う(/function ピボットの窓を開く\(おすすめか\)\{/.test(book),
  '★ピボットの窓を開く の 口は 戻してある（呼ぶ() が 渡している）★',
  '★消すと おすすめピボットが 効かなくなる★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  言う(見つけ.length > 上限 - 1,
    '★上限を 1つ 下げたら 引っかかる（今 ' + 見つけ.length + ' > ' + (上限 - 1) + '）★');
  const にせ = 見つけ.concat([{ ファイル: 'book.html', 行: 1, 名: 'startEdit',
    口の名: 'r, c, init', 口: 3, 最大: 2 }]);
  言う(にせ.filter((v) => v.名 === 'startEdit').length === 1,
    '★startEdit が 戻ったら ②が 落ちる★');
  言う(見つけ.filter((v) => v.名 === 'startEdit').length === 0, '★本物は 直ったまま★');
}

console.log('\nunused-param: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
