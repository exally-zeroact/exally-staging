/* shiki-tsunagi.test.mjs — ★自前の 土台に 自前の 関数を 繋ぐ 皮★（2026-09-14）
 *
 *  ★★何の 為の 皮か★★
 *    司さん 2026-09-13「借り物が あったら 商用など いろいろ 引っかかる。★自作で やれ★」
 *    借り物＝`hyperformula.full.min.js`（944,675バイト・★GPLv3★）
 *    ⇒ 中身（`lib/formula-*.js`）は もう 借り物から 独立して いる。
 *      ★借り物に ぶら下がって いたのは 皮（`lib/formula-*-plug.js`）だけ★
 *    ⇒ この 皮が その 差し替え＝★同じ 中身を 自前の 土台から 呼ぶ★
 *
 *  ★★この 見張りが 見る 物★★
 *    ①皮が 知って いる 名前の 数（★減ったら 赤★＝黙って 外れない）
 *    ②★土台だけで（借り物を 1行も 通さずに）皮の 関数が 答える★
 *    ③★四角（A1:B3）の 形（行数×列数）が 皮まで 届く★
 *      ＝届かないと ★横に 広がる 関数が 縦一列に 潰れる★（2026-09-14 実測）
 *    ④知らない 名前は ★#NAME?★（半分 合う 答えを 出さない）
 *    ⑤★どちらが 答えたかを 数えて いる★（借り物の 出番が 減る様子を 数で 見る為）
 *
 *  ★★見ていない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★溢れ（こぼれ）は まだ マスに 並べられません★（土台⑤）＝
 *      `=MUNIT(2)` は ★溢れの まま★ マスに 入る。★左上だけ 返して 誤魔化しては いません★
 *    ・答えが 実Excel と 同じかは ★別の 突き合わせ★で 見る
 *      （yosoku 721/724・nokori 83/83 ＝ `docs/measured/` の 紙）
 *    ・画面（book.html）には ★まだ 繋いで いません★
 *
 *  使い方: node tests/shiki-tsunagi.test.mjs
 *          node tests/shiki-tsunagi.test.mjs --self-test
 */
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const 皮 = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};
const 同じ = (得, 欲, なぜ) => {
  if (得 !== 欲) throw new Error((なぜ ? なぜ + ' … ' : '') + 'うち `' + 得 + '` ／ ★欲しい `' + 欲 + '`★');
};

/* ★この 数は 手で 決めた 物では ありません★＝
   `lib/formula-*-plug.js` の 名簿から 機械で 抜いて 繋いだ 数（extra13+kane22+filterxml1+
   cell1+yosoku10+nokori14 = 61）。★減ったら 黙って 外れた という 事★ */
/* 2026-09-16 … 61 → ★65★（BESSELI / BESSELJ / BESSELK / BESSELY を 足した）
     ★覚えた 係数表は 1つも 使って いません★ … lib/bessel.js
     ★実Excel の 方が 間違って いる 所が 3件 在ります★ … docs/measured/bessel-no-kotae.md */
/* 2026-09-19 ... 65 -> 67（ODDLPRICE / ODDLYIELD）-> ★69★（ODDFPRICE / ODDFYIELD ＝★保留 0個★）
     ＝経営者1 の 275本で 外れ 0本／わざと 壊したら 258本 赤
     ＝★台に 無い 27 → 25★ */
const 繋いだ数 = 69;

console.log('\n[shiki-tsunagi] ★自前の 土台に 自前の 関数を 繋ぐ★');

T('★皮が 知って いる 名前が ' + 繋いだ数 + '個 在る★（減ったら 黙って 外れて いる）', () => {
  同じ(皮.名前たち().length, 繋いだ数);
});

T('★★借り物を 1行も 通さずに 皮の 関数が 答える★★', () => {
  const h = H.表();
  h.打つ('A1', '=BAHTTEXT(1)');
  同じ(h.字('A1'), 'หนึ่งบาทถ้วน', 'BAHTTEXT（お金の タイ語）');
  h.打つ('A2', '=CONVERT(1,"m","cm")');
  同じ(h.字('A2'), '100', 'CONVERT（単位）');
  h.打つ('A3', '=ERF.PRECISE(0)');
  同じ(h.字('A3'), '0', 'ERF.PRECISE');
});

T('★土台の 7個は 今まで どおり★（皮を 足して 壊れて いない）', () => {
  const h = H.表();
  h.打つ('A1', '1'); h.打つ('A2', '2'); h.打つ('A3', '3');
  h.打つ('B1', '=SUM(A1:A3)');
  同じ(h.字('B1'), '6');
});

T('★★四角の 形（行数×列数）が 皮まで 届く★★（届かないと 縦一列に 潰れる）', () => {
  const h = H.表();
  /* A1:B3 = 1,2 / 3,4 / 5,6 */
  h.打つ('A1', '1'); h.打つ('B1', '2');
  h.打つ('A2', '3'); h.打つ('B2', '4');
  h.打つ('A3', '5'); h.打つ('B3', '6');
  h.打つ('D1', '=ARRAYTOTEXT(A1:B3,1)');
  /* ★実Excel の 実測★＝`{1,2;3,4;5,6}`。形が 届かないと `{1;2;3;4;5;6}` に なる */
  同じ(h.字('D1'), '{1,2;3,4;5,6}', '★横の 並びが 残って いるか★');
});

T('★四角の 形が 要る 行列も 通る★（MINVERSE は 正方で ないと #VALUE!）', () => {
  const h = H.表();
  h.打つ('A1', '4'); h.打つ('B1', '7');
  h.打つ('A2', '2'); h.打つ('B2', '6');
  h.打つ('D1', '=MINVERSE(A1:B2)');
  const v = h.値('D1');
  if (!v || v.溢れ !== true) throw new Error('溢れが 返って いない … ' + JSON.stringify(v));
  同じ(v.行数 + '行' + v.列数 + '列', '2行2列');
});

T('★知らない 名前は #NAME?★（半分 合う 答えを 出さない）', () => {
  const h = H.表();
  h.打つ('A1', '=KONNAKANSUUWANAI(1)');
  同じ(h.字('A1'), '#NAME?');
});

T('★★どちらが 答えたかを 数えて いる★★（借り物の 出番が 減る様子を 数で 見る）', () => {
  H.答えた数を消す();
  const h = H.表();
  h.打つ('A1', '1'); h.打つ('A2', '2');
  h.打つ('B1', '=SUM(A1:A2)');        /* 土台 */
  h.打つ('B2', '=BAHTTEXT(1)');        /* 皮 */
  h.打つ('B3', '=MUNIT(2)');           /* 皮（溢れ待ち） */
  /* ★知らない 式は 溢れの 外（B6）に 置く★（2026-10-05）
       前は B4 に 置いて いた ＝ B3 の `=MUNIT(2)` が 溢れる B3:C4 の ★中★。
       溢れた 後に 溢れの 中へ 打つと 実Excel は 元の 式を #スピル! に する（経営者の 物差し 137 の 組 M）＝
       台も そう 直した ので B3 が 2回 計算され「皮（溢れ待ち）」が 2 に なった。★この 試験が 見たいのは 数え方★なので 外へ 移す */
  h.打つ('B6', '=SHIRANAI(1)');        /* 知らない */
  /* ★読まないと 計算されません★＝この 土台は「読んだ 時に 計算する」作り。
     2026-09-14 に ここで 数が 0に なり ★見張りが 空振り★して いました。 */
  ['B1', 'B2', 'B3', 'B6'].forEach(function (k) { h.値(k); });
  const n = H.答えた数();
  同じ(n['土台'], 1, '土台');
  同じ(n['皮'], 1, '皮');
  同じ(n['皮（溢れ待ち）'], 1, '皮（溢れ待ち）');
  同じ(n['知らない'], 1, '知らない');
});

T('★★「今 いる マス」の 覚えが 入れ子で 漏れて いない★★（2026-09-15）', () => {
  /* ★訳★＝`=CELL("row")`（引数なし）は ★今 計算して いる マス★を 使います。
     土台は `計算する` の 中で 覚えて `finally` で 戻して います。
     ★戻し忘れると 別の マスの 番号を 答えます★＝
     ★指示役1 と 経営者1 の 案＝「同じ 式を 2回 押して 同じ 答えか」＋「間に 別のマスを 挟む」★ */
  const h = H.表();
  h.打つ('A3', '=CELL("row")');
  同じ(h.字('A3'), '3', '1回目（A3）');
  h.打つ('B7', '=CELL("row")');
  同じ(h.字('B7'), '7', '2回目（B7＝別の マス）');
  同じ(h.字('A3'), '3', '★3回目＝A3 に 戻るか（覚えが 漏れて いないか）★');
  /* ★互いに 読み合う 形★＝入れ子の 計算が 起きる */
  h.打つ('C5', '=CELL("row")+A3');
  同じ(h.字('C5'), '8', 'C5（5）＋A3（3）');
  同じ(h.字('A3'), '3', '★読み合った 後も A3 は 3★');
});

T('★★溢れは まだ マスに 並べられない と 書いて 在る★★（出来て いない事を 隠さない）', () => {
  const s = require_('node:fs').readFileSync(path.join(ROOT, 'lib/shiki-hyou.js'), 'utf8');
  if (!/溢れは まだ 置けません/.test(s)) throw new Error('★溢れが まだ な事が 書かれて いない★');
  if (!/左上だけ 返して 誤魔化す事は しません/.test(s)) throw new Error('★左上だけ 返さない と 書かれて いない★');
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-tsunagi --self-test] ★わざと 壊したら 赤に なるか★');
  T('★皮を 外したら 皮の 関数は #NAME? に なる★（見張りが 空振りして いない）', () => {
    const H2 = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
    /* 皮を 渡さずに 組み立て直す＝factory を 直に 呼ぶ 手が 無いので、
       ★皮の 表から 1つ 抜いて★ 同じ事を 見る */
    const 元 = 皮.表['BAHTTEXT'];
    delete 皮.表['BAHTTEXT'];
    try {
      const h = H2.表();
      h.打つ('A1', '=BAHTTEXT(1)');
      同じ(h.字('A1'), '#NAME?', '皮から 抜いた 後');
    } finally { 皮.表['BAHTTEXT'] = 元; }
  });
  T('★抜いた 物を 戻したら 緑に 戻る★', () => {
    const h = H.表();
    h.打つ('A1', '=BAHTTEXT(1)');
    同じ(h.字('A1'), 'หนึ่งบาทถ้วน');
  });
  T('★形を 潰したら ARRAYTOTEXT が 縦一列に なる★（形の 見張りが 効いて いる）', () => {
    const v = 皮.呼ぶ('ARRAYTOTEXT', [
      { 種: '四角', 並び: [1, 2, 3, 4, 5, 6].map((n) => ({ 型: '数', 値: n })) },
      { 種: '直', 値: { 型: '数', 値: 1 } },
    ]);
    同じ(String(v && v.値), '{1;2;3;4;5;6}', '★形を 持たせなければ 潰れる★');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
