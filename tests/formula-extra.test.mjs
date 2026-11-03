/* formula-extra.test.mjs — ★足りない 関数の 答えが 実Excel と 同じか★ 2026-08-31
 *
 *  ★なぜ★（司さん 2026-08-30）
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    実測（2026-08-31・実ブラウザ）＝エンジンに ★457個★。
 *    事務で 使う 98個を 見て ★62個 無かった★。うち ★12個★を 足した。
 *
 *  ★★答えは 実Excel を 打って 取った★★（2026-08-31・新規の 空ブック・保存なし）
 *      AVERAGEIFS(E1:E3,D1:D3,"A")   → 20
 *      AVERAGEIFS(E1:E3,D1:D3,"Z")   → #DIV/0!
 *      AVERAGEIFS(E1:E3,E1:E3,">=20")→ 25
 *      TAKE(A1:B3,2)                 → 1,2,3,4
 *      TAKE(A1:B3,-1)                → 5,6
 *      DROP(A1:B3,1)                 → 3,4,5,6
 *      DROP(A1:B3,0,1)               → 2,4,6
 *      CHOOSECOLS(A1:B3,2)           → 2,4,6
 *      CHOOSECOLS(A1:B3,3)           → #VALUE!
 *      CHOOSEROWS(A1:B3,-1)          → 5,6
 *      TOCOL(A1:B3)                  → 1,2,3,4,5,6
 *      TOROW({1,2;3,4},0,TRUE)       → 1,3,2,4
 *      WRAPROWS({1;2;3;4;5},2,0)     → 1,2,3,4,5,0
 *      EXPAND(A1:B3,4,3,0)           → 1,2,0,3,4,0,5,6,0,0,0,0
 *      EXPAND(A1:B3,2,2,0)           → #VALUE!（小さくは できない）
 *      ARRAYTOTEXT({1,2;3,4})        → 1, 2, 3, 4
 *      ARRAYTOTEXT({1,2;3,4},1)      → {1,2;3,4}
 *      MODE.MULT({1;2;2;3;3})        → 2,3
 *
 *  ★足しかけて 止めた 物★
 *    ★YEN は 実Excel に 無い★（実測＝#NAME?）。円は DOLLAR が 出す。
 *    ★無い 関数を 足すのは 捏造★なので 止めた。
 *
 *  ★形（spill）を 保つ★
 *    2026-08-29 に ★FILTER が 縦1列しか 返せず 月別合計が 全部 0★に なった。
 *    ⇒ ★2列以上でも 形（行数×列数）を 保つ★事を ここで 見る。
 *
 *  ★実ブラウザでも 打った★（2026-08-31）
 *    21通り 中 ★19通り 一致★。合わない 2つは ★`,,`（真ん中を 空ける）書き方★だけ
 *    ＝エンジン側の 決まり（lib/formula-extra-plug.js に 記録）。
 *
 *  走らせ方: node tests/formula-extra.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const F = require_(path.join(ROOT, 'lib/formula-extra.js'));

/* 材料（実Excel で 打った 時と 同じ） */
const 表 = [[1, 2], [3, 4], [5, 6]];
const 印 = [['A'], ['B'], ['A']];
const 値 = [[10], [20], [30]];

/** 平らにして カンマで つなぐ（TEXTJOIN と 同じ 見え方） */
function 平(v) {
  if (!Array.isArray(v)) return String(v && v.誤り ? '#' + v.誤り : v);
  const 出 = [];
  for (const r of v) for (const c of r) 出.push(c && c.誤り ? '#' + c.誤り : c);
  return 出.join(',');
}
const 誤り = (v) => (v && !Array.isArray(v) && v.誤り) ? v.誤り : null;

console.log('★足りない 関数（答えは 実Excel の 実測）★');

/* ── AVERAGEIFS ── */
言う(F.条件つき平均(値, [[印, 'A']]) === 20, '★AVERAGEIFS(E1:E3,D1:D3,"A") ＝ 20★');
言う(誤り(F.条件つき平均(値, [[印, 'Z']])) === 'DIV_BY_ZERO',
  '★当たりが 無ければ #DIV/0!★（実Excel と 同じ）');
言う(F.条件つき平均(値, [[値, '>=20']]) === 25, '★AVERAGEIFS(E1:E3,E1:E3,">=20") ＝ 25★');
言う(F.条件つき平均(値, [[印, 'a']]) === 20, '★大文字小文字を 区別しない★（実Excel と 同じ）');
言う(F.条件に合うか('売上円', '*円') === true, '★* が 使える★');
言う(F.条件に合うか('あい', '?い') === true, '★? が 使える★');
言う(F.条件に合うか(5, '<>5') === false, '★<> が 使える★');

/* ── TAKE / DROP ── */
言う(平(F.切り出す(表, 2, null, false)) === '1,2,3,4', '★TAKE(A1:B3,2) ＝ 1,2,3,4★');
言う(平(F.切り出す(表, -1, null, false)) === '5,6', '★TAKE(A1:B3,-1) ＝ 5,6★（尻から）');
言う(平(F.切り出す(表, 1, null, true)) === '3,4,5,6', '★DROP(A1:B3,1) ＝ 3,4,5,6★');
言う(平(F.切り出す(表, 0, 1, true)) === '2,4,6', '★DROP(A1:B3,0,1) ＝ 2,4,6★');
/* ★★#VALUE! では なく #CALC! でした★★（2026-09-15・実Excel の 実測 6本）
     `=DROP(A1:A5,5)` `=DROP(A1:A5,6)` `=DROP(A1:A5,45294)`
     `=TAKE(A1:A5,0)` `=TAKE(A1:A5,0.5)` `=TAKE(A1:A5,-0.5)`
   ★紙★ `docs/measured/kansuu46/golden-346-2026-09-08.tsv`
   ★★この 行は 実Excel に 聞かずに 書いて ありました★★
     ＝`tests/shiki-kansuu-kami.test.mjs`（紙 60枚で 押す 台）が 見つけました
   ⇒★`#CALC!` は「空の 表に なった」時／`#VALUE!` は「渡し方が 悪い」時★＝★別の 誤り★ */
言う(誤り(F.切り出す(表, 9, null, true)) === 'CALC', '★全部 落としたら #CALC!★（実Excel 実測）');

/* ── CHOOSECOLS / CHOOSEROWS ── */
言う(平(F.選び出す(表, [2], true)) === '2,4,6', '★CHOOSECOLS(A1:B3,2) ＝ 2,4,6★');
言う(誤り(F.選び出す(表, [3], true)) === 'VALUE', '★CHOOSECOLS(A1:B3,3) ＝ #VALUE!★');
言う(平(F.選び出す(表, [-1], false)) === '5,6', '★CHOOSEROWS(A1:B3,-1) ＝ 5,6★');
言う(誤り(F.選び出す(表, [0], true)) === 'VALUE', '★0 は #VALUE!★（1から 数える）');

/* ── TOCOL / TOROW ── */
言う(平(F.一本にする(表, 0, false, true)) === '1,2,3,4,5,6', '★TOCOL(A1:B3) ＝ 1,2,3,4,5,6★');
言う(平(F.一本にする([[1, 2], [3, 4]], 0, true, false)) === '1,3,2,4',
  '★TOROW({1,2;3,4},0,TRUE) ＝ 1,3,2,4★（列の順に 読む）');
言う(平(F.一本にする([[1, '', 2]], 1, false, true)) === '1,2', '★空を とばせる（2つ目の 引数＝1）★');

/* ★2つ目は『大きさ』では なく『無視する 印』＝0〜3 だけ★（2026-09-08 実測）
   紙 … docs/measured/kansuu46/golden-hikisuu-mihari-2026-09-08.tsv（15通り × 12引数）
     n=-2 -1 … #VALUE! ／ n=-0.5 0 0.5 1 1.5 2 2.5 3 3.5 … 通る ／ n=4 5 6 45294 … #VALUE!
   ⇒★直す前は 何を 渡しても 受けて いた＝`=TOCOL(A1:B5,DATE(2024,1,3))` が 答えを 返して いた★ */
言う(誤り(F.一本にする(表, 4, false, true)) === 'VALUE', '★無視の 印が 4 は #VALUE!★');
言う(誤り(F.一本にする(表, -1, false, true)) === 'VALUE', '★無視の 印が −1 は #VALUE!★');
言う(誤り(F.一本にする(表, 45294, false, true)) === 'VALUE',
  '★★日付を 渡すと #VALUE!★★（45294 ＝ 2024/1/3。直す前は 答えを 返して いた）');
言う(平(F.一本にする(表, 3.5, false, true)) === '1,2,3,4,5,6', '★3.5 は 切り捨てて 3＝通る★');
言う(平(F.一本にする(表, -0.5, false, true)) === '1,2,3,4,5,6',
  '★−0.5 は 切り捨てて 0＝通る★（★REPLACEB の 方は −0.5 を 断る＝関数ごとに 違う★）');

/* ── WRAPROWS / WRAPCOLS ── */
言う(平(F.折り返す([[1], [2], [3], [4], [5]], 2, 0, true)) === '1,2,3,4,5,0',
  '★WRAPROWS({1;2;3;4;5},2,0) ＝ 1,2,3,4,5,0★');
/* ★★ここは 元々 'VALUE' と 書いて ありましたが ★実Excel に 聞いて いませんでした★★
   実測（2026-09-08・同じ 紙）… =WRAPROWS(A1:A5,n) は n=-2 -1 -0.5 0 0.5 で ★#NUM!★
   ⇒★『試験が 緑』は『実Excel と 同じ』では ない＝試験の 方が 間違って いた★ */
言う(誤り(F.折り返す(表, 0, 0, true)) === 'NUM', '★幅0は ★#NUM!★（実Excel に 聞いた）');
言う(誤り(F.折り返す(表, -1, 0, true)) === 'NUM', '★幅−1も #NUM!★');
言う(誤り(F.折り返す(表, 0.5, 0, true)) === 'NUM', '★幅0.5も 切り捨てて 0＝#NUM!★');

/* ── EXPAND ── */
言う(平(F.広げる(表, 4, 3, 0)) === '1,2,0,3,4,0,5,6,0,0,0,0',
  '★EXPAND(A1:B3,4,3,0) ＝ 1,2,0,3,4,0,5,6,0,0,0,0★');
言う(誤り(F.広げる(表, 2, 2, 0)) === 'VALUE', '★小さくは できない（#VALUE!）★');

/* ── ARRAYTOTEXT ── */
言う(F.表を字に([[1, 2], [3, 4]], 0) === '1, 2, 3, 4', '★ARRAYTOTEXT({1,2;3,4}) ＝ 1, 2, 3, 4★');
言う(F.表を字に([[1, 2], [3, 4]], 1) === '{1,2;3,4}', '★ARRAYTOTEXT({1,2;3,4},1) ＝ {1,2;3,4}★');
言う(F.表を字に([['a'], [true]], 1) === '{"a";TRUE}',
  '★字は "" で 囲み・真偽は TRUE★（実Excel と 同じ）');

/* ── MODE.MULT ── */
言う(平(F.最頻値たち([[1], [2], [2], [3], [3]])) === '2,3', '★MODE.MULT({1;2;2;3;3}) ＝ 2,3★');
言う(誤り(F.最頻値たち([[1], [2], [3]])) === 'NA', '★1回ずつなら #N/A★（実Excel と 同じ）');

/* ── ★形（spill）を 保つ★ ── */
const t = F.切り出す(表, 2, null, false);
言う(Array.isArray(t) && t.length === 2 && t[0].length === 2,
  '★TAKE は 2行2列の 形を 保つ★（★2026-08-29 の FILTER の 失敗を 繰り返さない★）',
  JSON.stringify(t));
const col = F.一本にする(表, 0, false, true);
言う(col.length === 6 && col[0].length === 1, '★TOCOL は 縦1列（6行1列）★');
const row = F.一本にする(表, 0, false, false);
言う(row.length === 1 && row[0].length === 6, '★TOROW は 横1行（1行6列）★');
const mm = F.最頻値たち([[1], [2], [2], [3], [3]]);
言う(mm.length === 2 && mm[0].length === 1, '★MODE.MULT は 縦1列★（実Excel と 同じ）');

/* ── ★足さない 物に 理由★ ── */
const 数 = F.数える();
言う(数.足す.length === 12, '★足すのは 12個★（今 ' + 数.足す.length + '個）');
/* ★2026-09-04 に BAHTTEXT を 足した★＝★中身は lib/bahttext.js★（このファイルでは 計算しない）
   ⇒★繋ぐ 数は 12 + 1 = 13★／★数を ただ 上げるのでは なく「どこに 在るか」まで 見る★ */
言う(Object.keys(数.別のlibで足す).length === 1 && !!数.別のlibで足す.BAHTTEXT,
  '★別の lib で 足した 物は BAHTTEXT の 1つ★');
言う(/lib\/bahttext\.js/.test(数.別のlibで足す.BAHTTEXT), '★どこに 在るかが 書いてある★');
言う(!数.足さない.BAHTTEXT, '★足さない 一覧から BAHTTEXT が 消えている（もう 足した）★');
/* ★★2026-09-05 実測で 直した（08-31 の 決めつけが 試験に 焼いてあった）★★
   ★前★ … 「実Excel に 無い」「/存在しない/ と 書いてあるか」
   ★実測★ … `.FormulaLocal = "=YEN(1234)"` は ★¥1,234 を 出す★
             `.Formula`（英語の 構文）で 打つと #NAME?＝08-31 は ★こちらだけ 測っていた★
             ⇒★「無い」ではなく「日本語UIの 表示名で、本名は DOLLAR」★

   ★★もっと 悪かった 所★★
     YEN を ★「足さない」の 棚★に 置いていた
     ⇒ AIへ 渡る 紙(prompt/kansuu.md)で ★「Exally内では まだ 動かない」の 列★に 並ぶ
     ⇒★実際は ¥1,235 が 出るのに、AIが お客さんに「動きません」と 言う★所だった
     ⇒★★「足していない」と「動かない」は 別★★＝★別名で動く★の 棚を 足した

   ★本当に 動くかは text-format.test.mjs で ★engine に 通して★ 測る★
     （ここは 台帳の 字だけ＝それだけでは 半分） */
言う(数.足す.indexOf('YEN') < 0,
  '★YEN は 自分では 作らない（本名 DOLLAR に 直して 動かす＝別名を 増やさない）★');
言う(!数.足さない.YEN,
  '★★YEN が「動かない」の 棚に 居ない（居ると AIが 客に「動きません」と 言う）★★');
言う(!!数.別名で動く && !!数.別名で動く.YEN,
  '★YEN は「別名で動く」の 棚に 居る★', Object.keys(数.別名で動く || {}).join(' / '));
言う(/DOLLAR/.test(String((数.別名で動く || {}).YEN || '')),
  '★理由に「本名は 何か」が 書いてある★', (数.別名で動く || {}).YEN);
言う(/convertFormula/.test(String((数.別名で動く || {}).YEN || '')),
  '★理由に「どこが 直すか」が 書いてある★', (数.別名で動く || {}).YEN);
const 理由なし = Object.keys(数.足さない).filter((k) => !String(数.足さない[k]).trim())
  .concat(Object.keys(数.別名で動く || {}).filter((k) => !String(数.別名で動く[k]).trim()));
言う(理由なし.length === 0, '★足さない／別名で動く 物には 全部 理由が 書いてある★', 理由なし.join(' / '));
/* ══ ★★2026-09-19 ... ★LAMBDA と PHONETIC を 名指しで 決め打つのを やめました★★★ ══
     ★前★ ... 「LAMBDA と PHONETIC が ★足さない 棚に 在る★」を 門に して いた
     ★何が 起きたか★
       見張り（kansuu-tana / kansuu-kabaa）が ★2段しか 押して いなかった★ので
       ★本番では 動いて いる 物を 「動かない」棚に 置いて いました★
       ⇒2026-09-19 に 見張りを ★3段★（JS層 → 台 → 借り物）に 揃えた
       ⇒★LAMBDA / BYROW / BYCOL / PHONETIC が 4個とも 動くと 出た★（実測 8/8）
       ⇒★PHONETIC は 同じ日に 範囲の #VALUE! も 直した★
       ⇒★4個を 「別名で動く」棚へ 移した★
     ⇒★★「今 在る 物」を 名前で 決め打った 門は その 物が 動いた 日に 死にます★★
        ＝経営者1 の ㊲（誰も 打って いない）と 同じ 型
     ★★どう 書き直したか★★
       ★名前では なく ★決まり★を 見ます★
         ⑴どちらかの 棚に ★必ず 居る★（★黙って 消えて いない★）
         ⑵どの 棚に 居ても ★訳が 書いて ある★
       ⇒★棚を 移しても 赤に ならない／★消したら 赤に なる★★ */
{
  const 棚たち = [数.足さない, 数.別名で動く || {}, 数.形で動く || {}];
  const 居る = (n) => 棚たち.some((t) => Object.prototype.hasOwnProperty.call(t, n));
  const 訳 = (n) => {
    for (const t of 棚たち) if (Object.prototype.hasOwnProperty.call(t, n)) return String(t[n] || '').trim();
    return '';
  };
  const 見る = ['LAMBDA', 'PHONETIC', 'BYROW', 'BYCOL'];
  const 居ない = 見る.filter((n) => !居る(n));
  言う(居ない.length === 0,
    '★★LAMBDA・PHONETIC・BYROW・BYCOL は どれかの 棚に 必ず 居る★★（★黙って 消えて いない★）',
    居ない.join(' / '));
  const 訳なし = 見る.filter((n) => 居る(n) && !訳(n));
  言う(訳なし.length === 0, '★その 4つには 訳が 書いて ある★', 訳なし.join(' / '));
}

/* ── ★繋ぐ 側★ ── */
const plug = fs.readFileSync(path.join(ROOT, 'lib/formula-extra-plug.js'), 'utf8');
const { 注記を外す } = await import(
  pathToFileURL(path.join(ROOT, 'scripts/lib/chuki.mjs')).href);
const plug素 = 注記を外す(plug);
言う(/class extends/.test(plug素), '★extends で 作っている★',
  '★Object.create だと this.evaluateAst が undefined に なる（2026-08-31 実測）★');
言う(/onlyValues/.test(plug素), '★形を 保って 返している★');
言う(!/YEN/.test(plug素.replace(/\s/g, '')) || !/'YEN'/.test(plug素),
  '★繋ぐ 側にも YEN が 残っていない★');
const 実装 = (plug素.match(/'[A-Z.]+':\s*\{ method:/g) || []).length;
const 繋ぐ筈 = 数.足す.length + Object.keys(数.別のlibで足す).length;
言う(実装 === 繋ぐ筈, '★繋ぐ 数 ＝ このファイルの 12 ＋ 別の lib の '
  + Object.keys(数.別のlibで足す).length + ' ＝ ' + 繋ぐ筈 + '個★（今 ' + 実装 + '個）');
言う(/'BAHTTEXT':\s*\{ method: 'bahttext'/.test(plug素), '★BAHTTEXT を 繋いでいる★');
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
言う(/formula-extra\.js/.test(book) && /formula-extra-plug\.js/.test(book),
  '★画面が 両方 読み込んでいる★');
言う(/FormulaExtraPlug\.つなぐ\(HyperFormula/.test(注記を外す(book)),
  '★画面が エンジンに 繋いでいる★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  言う(誤り(F.条件つき平均([[]], [])) === 'DIV_BY_ZERO', '★空でも 落ちない★');
  言う(平(F.切り出す(表, null, null, false)) === '1,2,3,4,5,6', '★省くと 全部★');
  言う(F.表を字に([[null]], 0) === '', '★空は 空のまま★');
  言う(F.条件に合うか(null, null) === true, '★空どうしは 合う★');
  言う(F.条件に合うか(1, '>abc') === false, '★数と 字は 比べない★');
  /* ★実Excel でも 埋めを 省くと #N/A★（2026-08-31 実測）
       WRAPROWS({1;2;3},2)    → #N/A
       WRAPROWS({1;2;3},2,0)  → 1,2,3,0
     ★4つ目（行でか）を 渡さないと 列で 折り返す★ので ここでは true を 渡す。 */
  言う(平(F.折り返す([[1], [2], [3]], 2, undefined, true)) === '1,2,3,#NA',
    '★埋めを 省くと #N/A★（実Excel と 同じ）',
    平(F.折り返す([[1], [2], [3]], 2, undefined, true)));
  言う(平(F.折り返す([[1], [2], [3]], 2, 0, true)) === '1,2,3,0',
    '★埋めを 渡せば その 値★（実測＝1,2,3,0）');

  /* ★★本当に 壊して 赤に なるかを 見る★★
     ここまでは「端の 値を 通す」だけ＝★壊していない★。
     ★見出しで「壊す」と 名乗る 以上 本当に 壊す★（tests/name-vs-body.test.mjs の 決まり）。
     ★repo は 読むだけ★＝字の 上で 壊して 数え直す。 */
  const 元字 = fs.readFileSync(path.join(ROOT, 'lib/formula-extra-plug.js'), 'utf8');
  /* ★印は 実物どおりに★＝間に 空白が 並んでいる（そろえて 書いてある）。
     ★合わない 印で 置き換えると 何も 壊れず「壊した つもり」に なる★ */
  const 壊した = 元字.replace(/'TAKE':\s*\{ method:/, "'TAKE': { nope:");
  言う(壊した !== 元字, '★壊せた（TAKE を 別名に した）★');
  const 数え = (t) => (注記を外す(t).match(/'[A-Z.]+':\s*\{ method:/g) || []).length;
  /* ★2026-09-04 に 13個に なった★（BAHTTEXT を 足した＝中身は lib/bahttext.js）
     ⇒★数を 決め打ちに しない★＝★1つ 消したら 1つ 減る事★を 見る（数える所が 効いている） */
  言う(数え(壊した) === 数え(元字) - 1,
    '★1つ 消したら 1つ 減る（数える所が 効いている）★',
    '本物 ' + 数え(元字) + ' / 壊した ' + 数え(壊した));
  言う(数え(元字) === 数.足す.length + Object.keys(数.別のlibで足す).length,
    '★繋いでいる 数が 台帳と 合う★（今 ' + 数え(元字) + '個）');
  const 表壊し = 元字.replace('class extends', 'class Nope');
  言う(!/class extends/.test(表壊し),
    '★extends を 外したら ③が 落ちる★');
  言う(fs.readFileSync(path.join(ROOT, 'lib/formula-extra-plug.js'), 'utf8') === 元字,
    '★★本物の ファイルは 1バイトも 触っていない★★');
}

console.log('\nformula-extra: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
