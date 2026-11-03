/* marume-mihari.test.mjs — ★答えを 丸めて いる 所が 増えて いないか★
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★一覧に 無い 丸めが 増えたら 赤★
 *      ⇒★数だけ 見ない★（25という 数は 行が 動けば ずれる）
 *      ⇒★ファイル:行:字 の 一覧と 突き合わせる★
 *    ②★A群（お客さんの 計算の 答えを 黙って 丸める 物）が 0 か★
 *      ⇒ 復活したら 赤
 *    ③★B・C・D・E 群を 理由つきで 一覧に 残す★（★消さない・混ぜない★）
 *
 *  ★★★A群かどうかを 機械に 判じさせない★★★（経営者・監査の 念押し）
 *    A（客の 答え）と C（見た目）は ★字の 形が まったく 同じ★
 *      A … `Math.round(_jsXirr(…)*10000)/10000`      ← 客の 答え
 *      C … `Math.round(v2 * 100) / 100`               ← グラフの 目盛りの 字
 *    ⇒★機械には 分けられません★
 *    ⇒★★だから 機械は「一覧に 無い 丸めが 出たか」だけを 見る★★
 *    ⇒★新しい 丸めは ★全部 赤★／★人が A〜E に 分けて 一覧に 足す★
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★借り物は 見て いません★ … `hyperformula.full.min.js` ／ `lib/*.min.js`
 *      （★中身を 読むと『それを 元に 作った 物』と 言われる 筋が 生まれる★）
 *    ・`tests/` の 中は 見て いません（試験の 中の 丸めは 別の 話）
 *    ・`scripts/` `docs/` の 中は 見て いません
 *    ・★丸め以外の「答えを 変える 事」（切り捨て・桁上げ・型の 変換）は 見て いません★
 *
 *  ★どうやって 一覧を 作ったか★
 *    経営者（監査）の 条件A「★探す前に 範囲を 決めるな★」
 *    ⇒ 丸めの 書き方を ★7通り 先に 書き出してから★ 探した（下の 型）
 *    ⇒★★1回目は 25か所と 数えたが、探す 形を 直したら 27か所★★（XIRR を 数え落として いた）
 *    ⇒ 1つずつ 目で 見て A〜E に 分けた
 *
 *  使い方: node tests/marume-mihari.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ══ ★探す 形（★決め打ちを 避ける 為に 先に 7通り 書き出した★）★ ══ */
/* ★★`[^)]*` では 中に 括弧の 在る 物を 拾えない★★（2026-09-08 に 踏んだ）
     `Math.round(_jsXirr(vals,dates)*10000)/10000`
       ⇒ `[^)]*` は `_jsXirr(vals,dates` の 閉じ括弧で 止まる ⇒★拾えない★
     ⇒★私の 1回目の「25か所」は ★XIRR を 数え落として いた★★
     ⇒★★一番 直したい 物が 一覧から 漏れて いた＝見張りが 守る つもりの 物を 見て いなかった★★
   ⇒ `(?:[^()]|\([^()]*\))*` に 直した＝★1段の 入れ子まで 跨げる★

   ★★2回目の 見落とし（2026-09-08・A群の 丸めを 外しに 行って 見つけた）★★
     `Math.round(_jsBinomDistRange(parseInt(mBdr[1]),…)*10000)/10000`
       ⇒★入れ子が 2段★＝1段用の 形では 拾えない
       ⇒★コードに 在るのに 一覧に 無く、見張りは ★緑の まま★だった★
     ⇒★同じ 型の 見落としが 2回 続いた＝★正規表現を 広げる 直し方が 間違い★★
   ⇒★★中身を 正規表現で 読むのを やめた＝★括弧を 数えて 閉じを 探す★★★
     ★これで 入れ子の 段数に 上限が 無くなった★ */
function 閉じを探す(s, 開き) {
  let 深さ = 0;
  for (let i = 開き; i < s.length; i++) {
    if (s[i] === '(') 深さ++;
    else if (s[i] === ')') { 深さ--; if (深さ === 0) return i; }
  }
  return -1;   /* ★閉じて いない★ */
}

const 桁 = '(?:10{2,}|1e\\d+)';
const 型 = [
  /* ★頭★＝呼び出しの 始まり ／ ★中★＝括弧の 中身に 要る 形 ／ ★尻★＝閉じ括弧の 後ろに 要る 形 */
  { 札: '①Math.round(x*10^n)/10^n', 頭: /Math\.round\s*\(/g,
    中: new RegExp('\\*\\s*' + 桁 + '\\s*$'), 尻: new RegExp('^\\s*\\/\\s*' + 桁) },
  { 札: '②Math.round(x*10)/10', 頭: /Math\.round\s*\(/g,
    中: /\*\s*10\s*$/, 尻: /^\s*\/\s*10\b/ },
  { 札: '③toFixed', 式: /\.toFixed\s*\(\s*\d+\s*\)/g },
  { 札: '④toPrecision', 式: /\.toPrecision\s*\(\s*\d+\s*\)/g },
  { 札: '⑤floor/ceil に 桁', 頭: /Math\.(?:floor|ceil)\s*\(/g,
    中: new RegExp('\\*\\s*' + 桁 + '\\s*$'), 尻: new RegExp('^\\s*\\/\\s*' + 桁) },
  { 札: '⑥parseFloat(…toFixed)', 頭: /parseFloat\s*\(/g, 中: /\.to(?:Fixed|Precision)\s*\(/ },
  { 札: '⑦Number(…toFixed)', 頭: /Number\s*\(/g, 中: /\.to(?:Fixed|Precision)\s*\(/ },
  /* ★⑧を 足した（2026-09-11）★＝`toExponential` も ★桁を 決めて 丸める★ 物
     ★見張りが 見て いなかった★＝本番の コードに 2か所 在ったのに 一覧に 1つも 無かった
     （`lib/shiki-keisan.js` と `book.html`。★見つけたのは 新しい 台を 足した 日★） */
  { 札: '⑧toExponential', 式: /\.toExponential\s*\(\s*[^)]*\)/g },
];

/* ══ ★★一覧（★人が 1つずつ 目で 見て 分けた★）★★ ══════════════
   ★鍵は ファイル＋当たった 字★（★行の 番号は 使わない＝行が 動いても ずれない★）
   ★群★
     A … ★お客さんの「計算の 答え」を 黙って 丸めて いる★（★0 で ないと 赤★）
     B … 実測で「そうする」と 決めた 所（触らない）
     C … 画面に 出す 見た目だけ（計算の 答えでは ない）
     D … ★丸めた と 分かる ように して 在る（正しい 形）★
     E … 判じかねる（★測ってから★＝別件） */
const 一覧 = [
  /* ★A群★ … `docs/measured/golden-marume-A-2026-09-08.tsv` で 実Excel と 突き合わせ済み */
  { 道: 'exally-formula.js', 字: 'Math.round(_jsXirr(vals,dates)*10000)/10000', 群: 'A',
    訳: 'XIRR … 実Excel 0.5478870809078217 が 0.5479 に なる（相対 2.4e-5）' },
  { 道: 'exally-formula.js', 字: 'Math.round(r[0]*10000)/10000', 群: 'A',
    訳: 'LINEST … 実Excel 0.00108 が 0.0011 に なる（相対 1.85e-2＝約1.9%）' },
  { 道: 'book.html', 字: 'Math.round(result*10000)/10000', 群: 'A',
    訳: '数式バーの 逃がし道（evalFormula）… 答えを 黙って 4桁に して いた'
      + '／★★お客さんが 踏む 形を 私は まだ 1つも 出せて いない★★'
      + '＝呼ぶ 所 3つとも ★エンジンが 失敗した 時だけ★（book.html:2555 は 誤りの 時／:9588・:9669 は setCellFormula が null の 時）'
      + '／★前に 私は「=1/3 が 0.3333 に なる」と 書いたが、それは ★押して 確かめて いない★（=1/3 は エンジンが 答えるので この 道を 通らない）★' },
  /* ★★4か所目★★（2026-09-08 に 見つけた＝★見張りが 入れ子 2段を 読めず 見えて いなかった★）
     `docs/measured/golden-marume-A4-binom-2026-09-08.tsv` で 実Excel と 突き合わせ済み
     ★12本 測って 3本が 4桁で ★ちょうど 0★ に なった＝★起きる事が 起きない事に 変わる★ */
  { 道: 'exally-formula.js',
    字: 'Math.round(_jsBinomDistRange(parseInt(mBdr[1]),parseFloat(mBdr[2]),parseInt(mBdr[3]),mBdr[4]?parseInt(mBdr[4]):undefined)*10000)/10000',
    群: 'A',
    訳: 'BINOM.DIST.RANGE … 実Excel 2.1426377248779138E-07 が ★0★ に なる（相対 1.0）／ずれ 最大 は =…(1000,0.5,550) の 1.807E-1' },

  /* ★B群★ … 実測で「そうする」と 決めた 所＝★触らない★ */
  { 道: 'lib/formula-complex.js', 字: '.toPrecision(15)', 群: 'B',
    訳: 'IM系 … 2026-09-08 に 78本を これで 直した（実Excel と 同じ 15桁）' },
  { 道: 'book.html', 字: '.toPrecision(15)', 群: 'B',
    訳: 'forDisplay … 「Excel も 15桁で 見せる」と 実測が 書いて 在る' },
  { 道: 'lib/bahttext.js', 字: '.toPrecision(15)', 群: 'B',
    訳: 'タイ語の 金額 … 116通り 実測ずみ' },
  /* ★★2026-09-16 に 足した 2つ★★ … `lib/shiki-kansuu.js` の ★複素数を 字に する 1行★
     ★訳★ … ★複素数は ★字を 返す 関数★★＝★その 字が そのまま マスに 出ます★
       実Excel は 15桁に 丸めて 見せます（実測）
         `=IMDIV(2,3)` → ★0.666666666666667★（★15桁★）
     ⇒★★`lib/formula-complex.js` の B群と ★同じ 訳★★★
       ＝借り物を 継いで いた 所を 自前で 書き直した（2026-09-16・棚㊼）ので
         ★同じ 丸めが 台の 側にも 要りました★
     ★A群では ありません★＝★数を 返す 関数（IMABS IMREAL …）には 掛けて いません★
       （掛けると 壊れる … `=IMARGUMENT(-2)` は 実Excel が ★17桁★で 返す・棚に 在る） */
  { 道: 'lib/shiki-kansuu.js', 字: '.toPrecision(15)', 群: 'B',
    訳: '複素数を 字に する 所 … 実Excel も 15桁で 見せる（=IMDIV(2,3) → 0.666666666666667）' },
  { 道: 'lib/shiki-kansuu.js', 字: 'Number(n.toPrecision(15))', 群: 'B',
    訳: '同じ 1行（型が 違うので 別に 当たる）… 字を 返す IM系 だけ／数を 返す 物には 掛けない' },
  /* ★2026-09-15 に 足した 2つ★ … `lib/shoshiki.js`（★書式の 台★）の 同じ 1行
     ★実Excel は 書式を 掛ける 前に 15桁に 丸めます★（2026-09-11 に 実物で 見つけた 決まり）
       34779.49999999999 … 実Excel ★34,780★／丸めないと ★34,779★＝★お金が 1円 ずれる★
     ★前は 呼ぶ側（`book.html:3440` の `_十五桁`）が やって いました★
       ⇒★`js/book-open.js:378` には 入って いなかった★＝★道が 2本／片方だけ 直って いた★
       ⇒★台の 中に 入れた★＝★呼ぶ側が 忘れられません★
     ★B群です★＝★実測で「そうする」と 決めた 所★（C群＝見た目だけ、では ない。
       ★TEXT() の 答えにも 効く＝式の 答えの 字に なる★ので B群に 置きます）
     ★確かめ★ … 司さんの 実物 ★16,536マスで 今の 画面と 1マスも 違わない★
                 ／紙 192行と 1行も 違わない／`kumikomi-shoshiki`（1円ずれの 見張り）緑 */
  { 道: 'lib/shoshiki.js', 字: '.toPrecision(15)', 群: 'B',
    訳: '書式の 台 … ★書式を 掛ける 前に 15桁★（34779.49999999999 → 34,780。丸めないと 1円 ずれる）' },
  /* ★★DB（定率法）の 3桁丸め★★（2026-09-16・棚61）
     ★実Excel 自身が ★値下がり率を 3桁に 丸めて★ います（文書に も 書いて ある）
     ★実測で 確かめました★
       `=DB(D1,D2,0.05,0.1,2)` … 実Excel ★-2845.8473333333332★
       これを 元×月/12（＝7548.66667）で 割ると ★ちょうど -0.377★
       ＝★丸めない 値（-0.377425…）では 合いません★
     ★A群では ありません★＝★客の 答えを 黙って 丸めて いるのでは なく
       ★丸めないと 実Excel と 違う★＝★真似る 側の 決め★ */
  { 道: 'lib/shiki-kansuu.js', 字: 'Math.round(割 * 1000) / 1000', 群: 'B',
    訳: 'DB（定率法）… ★実Excel が 値下がり率を 3桁に 丸める★（=DB(D1,D2,0.05,0.1,2) → -2845.8473333333332 は 割合 ちょうど -0.377）' },
  { 道: 'lib/shoshiki.js', 字: 'Number(n.toPrecision(15))', 群: 'B',
    訳: '同じ 1行（★型が 違うので 別に 当たる★＝`book.html` の forDisplay と 同じ 形）' },

  /* ★C群★ … 画面に 出す 見た目だけ（計算の 答えでは ない） */
  { 道: 'lib/chart.js', 字: 'Math.round(v2 * 100) / 100', 群: 'C', 訳: 'グラフの 目盛りの 字' },
  { 道: 'lib/chart.js', 字: 'Math.round(xv * 100) / 100', 群: 'C', 訳: 'グラフの 横軸の 字' },
  { 道: 'lib/chart.js', 字: 'Math.round(v * 100) / 100', 群: 'C', 訳: 'グラフの 縦軸の 字' },
  { 道: 'lib/chart.js', 字: 'Math.round(n * 100) / 100', 群: 'C', 訳: 'グラフの 数の 字' },
  { 道: 'lib/ink.js', 字: 'Math.round(mm * 一mmの点 * 100) / 100', 群: 'C', 訳: 'mm → 点（描画）' },
  /* ★2026-09-10 に 足した 2つ★ … `forDisplay`（★画面に 出す 字だけ★・計算は 触らない）
     ★実Excel の General は 決まった 字数に 収める★ことを 実測して 入れた
       `docs/measured/golden-shisuu-mitame-2026-09-10.tsv`（41の 数 × 14の 幅 ＝ 574本）
     ★上の B群に 同じ 行の `.toPrecision(15)` が 在ります★（型が 違うので 別に 当たる）
     ⇒★消して いません＝両方 残して あります★ */
  { 道: 'book.html', 字: 'Number(n.toPrecision(15))', 群: 'C',
    訳: 'forDisplay … ★実Excel は まず 15桁に 丸めてから 字数に 収める★（26718.499999999996 が 幅5で ★26719★＝26718 では ない）' },
  /* ★2026-09-11 に 足した★ … `_十進を幅に`（★画面に 出す 字だけ★・計算は 触らない）
     ★実Excel は 列の 幅で 出る 字が 変わります★（574本 実測）
       =1/3 … 幅40点 ★0.333★／幅80点 ★0.333333333★
     ⇒★十進の 形の まま 列の 幅に 収める為の 丸め★（桁数は ★幅から 出す＝決め打ちで ない★） */
  { 道: 'book.html', 字: 'Number(a.toPrecision(d))', 群: 'C',
    訳: '_十進を幅に … ★列の 幅に 収める★（幅40点 0.333／幅80点 0.333333333）／★計算は 触らない★' },
  { 道: 'book.html', 字: 'Number(a.toPrecision(d十))', 群: 'C',
    訳: 'forDisplay … ★十進で 出す 時の 有効桁★（桁数は ★字数から 出す★＝決め打ちで ない）／=1/3 が 実Excel と 同じ 0.333333333 に なる' },
  { 道: 'lib/review.js', 字: '.toFixed(1)', 群: 'C', 訳: '色の 明るさの 比の 字' },
  { 道: 'book.html', 字: 'Math.round(scale*10)/10', 群: 'C', 訳: '拡大率の 表示' },
  { 道: 'book.html', 字: '.toFixed(2)', 群: 'C', 訳: 'TEXT の 書式 0.00 / 0.00%（★書式は 丸めるのが 正しい★）' },
  { 道: 'book.html', 字: 'Math.round(r.傾き * 1e6) / 1e6', 群: 'C', 訳: '予測シートの 説明の 字' },
  { 道: 'book.html', 字: 'Math.round(r.切片 * 1e6) / 1e6', 群: 'C', 訳: '予測シートの 説明の 字' },
  { 道: 'book.html', 字: 'Math.round(r.傾き * 100) / 100', 群: 'C', 訳: '予測シートの 知らせの 字' },
  { 道: 'book.html', 字: 'Math.round(lastN*1e12)/1e12', 群: 'C', 訳: '連続データの 刻み' },

  /* ★D群★ … ★丸めた と 分かる ように して 在る（正しい 形）★ */
  { 道: 'lib/grid-stats.js', 字: 'Math.round(n * 1e10) / 1e10', 群: 'D',
    訳: '帯の 平均 … ★丸めた 時だけ ≒ を 付ける（黙って 丸めない）★' },
  { 道: 'lib/grid-stats.js', 字: 'Math.round(n * 10000) / 10000', 群: 'D',
    訳: '帯の 平均 … 同上（`(r===n ? "" : "≒") + fmt(r)`）' },

  /* ★★2026-09-11 に 足した 分★★（土台を 自分で 作る ④枚目＝`lib/shiki-keisan.js`）
     ★この 台の 丸めは 2か所だけ★＝どちらも ★画面に 出す 字を 作る 所★。
     計算の 途中（＋−×÷＾）では ★1回も 丸めて いません★
       ＋と − は 丸めるのでは なく ★消え残りを 0に する★（実Excel を 測って 決めた）
     ★実Excel と 突き合わせ済み★ … tests/shiki-keisan.test.mjs の 162通り */
  { 道: 'lib/shiki-keisan.js', 字: '.toPrecision(15)', 群: 'B',
    訳: '`数を字に`（平ら）… ★実Excel の 標準は 15桁で 丸める★（`=2/3&""` → 0.666666666666667）' },
  { 道: 'lib/shiki-keisan.js', 字: '.toExponential(14)', 群: 'B',
    訳: '`数を字に`（指数）… ★実Excel の 指数の 字は 15桁★（`=1.23456789012345E+21&""` そのまま）' },
  { 道: 'book.html', 字: '.toExponential(d - 1)', 群: 'C',
    訳: '★列の 幅に 収める為の 指数の 字★（計算の 答えでは ない）／★⑧を 足して 初めて 見えた★' },

  /* ★E群★ … 判じかねる（★測ってから★＝別件） */
  { 道: 'lib/formula-nokori.js', 字: 'Math.round(x * 1e10) / 1e10', 群: 'E',
    訳: '逆行列 … 「丸めの ごみを 落とす」と 書いて 在るが ★10桁は 15桁より 粗い★／実Excel と 未突合' },
  { 道: 'lib/formula-nokori.js', 字: 'Math.round(出 * 1e10) / 1e10', 群: 'E',
    訳: '順位の割合 … 同上' },
  { 道: 'book.html', 字: 'Math.round(答え * 1e6) / 1e6', 群: 'E',
    訳: 'ゴールシーク … きれいな 数に 寄せる。実Excel と 未突合' },
  { 道: 'book.html', 字: 'Math.round((線.傾き * x + 線.切片) * 1e6) / 1e6', 群: 'E',
    訳: '予測シート … ★マスに 書き込む 数★（見た目の 字では ない）。実Excel の 予測シートと 未突合' },

  /* ★★1回目の 探し方では 拾えて いなかった 2つ★★（2026-09-08）
     ①`Math.round(_jsXirr(vals,dates)*10000)/10000` … ★中に 括弧が 在る★
       ⇒ `[^)]*` が `_jsXirr(vals,dates` の 閉じ括弧で 止まる ⇒★一番 直したい 物が 漏れて いた★
     ②`Number(n.toPrecision(15))` … ★同じ 1行が ④と ⑦の 2つの 型に 当たる★
       ⇒ 1回目は ④だけ 数えて ⑦を 一覧に 入れ忘れた
     ⇒★探す 形を 直したら 25 → 27か所に なった★ */
  { 道: 'lib/formula-complex.js', 字: 'Number(n.toPrecision(15))', 群: 'B',
    訳: 'IM系 … ★上の toPrecision(15) と 同じ 1行★（④と ⑦の 両方に 当たる）。触らない'
      + '／★字は 括弧を 数えて 読むように 直した 分 伸びた（前は `Number(n.toPrecision(` で 切れて いた）★' },
];

/* ══ ★見る 範囲（★先に 数えて 書く★）★ ══════════════════════ */
function 見るファイル() {
  const 出 = [path.join(ROOT, 'exally-formula.js'), path.join(ROOT, 'book.html')];
  for (const f of fs.readdirSync(path.join(ROOT, 'lib'))) {
    if (!f.endsWith('.js') || f.endsWith('.min.js')) continue;   /* ★借り物は 見ない★ */
    出.push(path.join(ROOT, 'lib', f));
  }
  const k = path.join(ROOT, 'lib/keisan');
  if (fs.existsSync(k)) for (const f of fs.readdirSync(k)) if (f.endsWith('.js')) 出.push(path.join(k, f));
  return 出;
}

function 探す(読む) {
  const 出 = [];
  for (const p of 見るファイル()) {
    const s = 読む(p);
    if (s === null) continue;
    const 行ら = s.split('\n');
    const 積む = (始, 終, 札) => {
      const 行 = s.slice(0, 始).split('\n').length;
      出.push({
        道: path.relative(ROOT, p).replace(/\\/g, '/'),
        行, 型: 札, 字: s.slice(始, 終).replace(/\s+/g, ' ').trim(),
        文: String(行ら[行 - 1] || '').trim(),
      });
    };
    for (const t of 型) {
      if (t.式) {                       /* ★括弧を 跨がない 形（③④）＝そのまま★ */
        t.式.lastIndex = 0;
        let m;
        while ((m = t.式.exec(s)) !== null) 積む(m.index, m.index + m[0].length, t.札);
        continue;
      }
      /* ★★括弧を 数えて 読む（①②⑤⑥⑦）＝入れ子 何段でも 落ちない★★ */
      t.頭.lastIndex = 0;
      let m;
      while ((m = t.頭.exec(s)) !== null) {
        const 開き = m.index + m[0].length - 1;      /* 頭の 最後の 文字が '(' */
        const 閉じ = 閉じを探す(s, 開き);
        if (閉じ < 0) continue;                       /* ★閉じて いない＝数えない★ */
        const 中 = s.slice(開き + 1, 閉じ);
        if (t.中 && !t.中.test(中)) continue;
        if (t.尻) {
          const 後 = s.slice(閉じ + 1, 閉じ + 24);
          if (!t.尻.test(後)) continue;
          const 尻m = 後.match(t.尻);
          積む(m.index, 閉じ + 1 + 尻m[0].length, t.札);
        } else {
          積む(m.index, 閉じ + 1, t.札);
        }
      }
    }
  }
  return 出;
}

const 素で読む = (p) => (fs.existsSync(p) ? fs.readFileSync(p, 'utf-8') : null);

/* ★一覧の 鍵★＝ファイル＋字（空白を 詰めて 比べる） */
const 鍵 = (x) => x.道 + '|' + String(x.字).replace(/\s+/g, ' ').trim();
const 一覧の鍵 = new Set(一覧.map(鍵));

console.log('\n★丸めの 見張り★');
console.log('  ★見る 範囲★ … ' + 見るファイル().length + '本'
  + '（★借り物の min.js／tests／scripts／docs は 見て いません★）');

T('★探す 形が 8通り 在る（★決め打ちを 避けた 証拠★）★', () => {
  if (型.length !== 8) throw new Error('★型が ' + 型.length + '通り★（8通りのはず）');
});

/* ★★数を 直した（2026-09-08）★★
     題は「27か所」と 書いて あったが ★一覧は 25個★だった＝★題と 中身が 食い違って いた★
     ⇒ 括弧を 数えて 読むように 直し、A群の 4か所目（BINOM）を 足して ★一覧 26個★
     ⇒★題の 数と 期待の 数を 一覧から 引き直した★ */
/* ★★数を 直した（2026-09-10）★★
     `forDisplay` を ★実Excel の General（決まった 字数に 収める）★に した ので
     ★C群が 2つ 増えた★（12 → ★14★・一覧 26 → ★28★）
     ⇒★前の 数は 消して いません＝上に そのまま 残して あります★ */
T('★一覧が 痩せて いない（★34か所★・A4/B8/C16/D2/E4）★', () => {
  const ご = {};
  for (const x of 一覧) ご[x.群] = (ご[x.群] || 0) + 1;
  /* ★2026-09-11 … 幅で 出る 字を 変える 直しで ★C群が 1つ 増えた★（14 → 15・一覧 28 → 29）★
     ★前の 数は 消して いません＝上に そのまま 残して あります★ */
  /* ★2026-09-11（その2）… ★探す 形に ⑧toExponential を 足した★ ので
       ★前から 本番に 在ったのに 見えて いなかった 丸めが 1か所 見つかった★（book.html）
       ＋ 土台の 新しい 台（lib/shiki-keisan.js）の 2か所
       ⇒ B群 4 → ★6★ ／ C群 15 → ★16★ ／ 一覧 29 → ★32★ */
  /* ★★2026-09-15 … 書式の 台（`lib/shoshiki.js`）を 作ったので B群が 2つ 増えた★★
       （6 → ★8★・一覧 32 → ★34★）★前の 数は 消して いません＝上に そのまま★
     ★中身★ … `.toPrecision(15)` と `Number(n.toPrecision(15))`＝★同じ 1行★（型が 違うので 別に 当たる）
     ★訳★ … ★実Excel は 書式を 掛ける 前に 15桁に 丸める★（34779.49999999999 → 34,780）
             ★丸めないと お金が 1円 ずれます★（2026-09-11 に 実物で 見つけた 決まり）
             ★前は 呼ぶ側が やって いて `js/book-open.js:378` には 入って いなかった★
             ⇒★台の 中に 入れた＝呼ぶ側が 忘れられない★
     ★B群に した 訳★ … ★見た目だけ（C群）では ありません★＝
             ★TEXT() の 答えの 字にも 効く★＝★式の 答えに なる★ので B群 */
  /* ★★数は 決め打ちです★★＝★一覧を 黙って 痩せさせない 為★
       ★足した 時は ここも 直す★（★直さないと 赤で 止まる＝それが 狙い★）
       2026-09-16 … B群 8 → ★10★（`lib/shiki-kansuu.js` の 複素数を 字に する 1行）
         ＝借り物を 継いで いた IM系 26個を 自前で 書いた 分（棚㊼） */
  /* 2026-09-16 … B群 10 → ★11★（`lib/shiki-kansuu.js` の DB の 3桁丸め・棚61）
       ＝★実Excel 自身が 値下がり率を 3桁に 丸める★＝真似る 側の 決め */
  const 期待 = { A: 4, B: 11, C: 16, D: 2, E: 4 };
  const 出 = Object.keys(期待).map((k) => k + (ご[k] || 0)).join(' ');
  for (const k of Object.keys(期待)) {
    if ((ご[k] || 0) !== 期待[k]) throw new Error('★' + k + '群が ' + (ご[k] || 0) + '個★（' + 期待[k] + '個のはず）… 今 ' + 出);
  }
  console.log('      A' + ご.A + ' B' + ご.B + ' C' + ご.C + ' D' + ご.D + ' E' + ご.E + ' ＝ ' + 一覧.length + 'か所');
});

T('★一覧の 全部に 理由が 書いて ある（★消さない・混ぜない★）★', () => {
  const なし = 一覧.filter((x) => !x.訳 || !x.訳.trim());
  if (なし.length) throw new Error('★' + なし.length + '個に 理由が 無い★');
});

T('★★一覧に 無い 丸めが 増えて いない★★', () => {
  const 今 = 探す(素で読む);
  const 新しい = 今.filter((x) => !一覧の鍵.has(鍵(x)));
  console.log('      今 見つかった … ' + 今.length + 'か所 ／ 一覧 ' + 一覧.length + 'か所');
  if (新しい.length) {
    throw new Error('★一覧に 無い 丸めが ' + 新しい.length + 'か所★\n'
      + '      ⇒★機械は A群かを 判じません。★人が 見て A〜E に 分けて 一覧に 足して ください★\n'
      + 新しい.map((x) => '      ' + x.道 + ':' + x.行 + '  ' + x.字 + '\n          ' + x.文.slice(0, 110)).join('\n'));
  }
});

/* ★★2026-09-08 に 締めた★★
     前は「増えて いなければ 緑」だった＝★4か所 残った ままでも 緑★
     ⇒ A群を 4か所とも 外したので ★コードに 1つでも 残って いたら 赤★に する
     ⇒★一覧の 4個は 消さずに 残す★＝「昔 在って 外した」の 証拠（★消さない・混ぜない★） */
T('★★A群が コードに 1つも 残って いない（お客さんの 答えを 黙って 丸める 物 0か所）★★', () => {
  const A = 一覧.filter((x) => x.群 === 'A');
  const 今 = 探す(素で読む);
  const 残り = A.filter((a) => 今.some((x) => 鍵(x) === 鍵(a)));
  console.log('      A群 … 一覧 ' + A.length + '個（外した 記録）／ まだ コードに 在る ' + 残り.length + '個');
  if (残り.length) {
    throw new Error('★A群が ' + 残り.length + 'か所 コードに 残って いる★\n'
      + 残り.map((x) => '      ' + x.道 + '  ' + x.字).join('\n'));
  }
});

/* ★A群と 同じ 形（*10000)/10000）が どこにも 復活して いない事も 見る★
   ⇒★一覧に 足せば 通る 抜け道★を 塞ぐ＝★黙って 4桁に する 形は もう 使わない★

   ★★免除は 1か所だけ（★理由つき★・黙って 見逃さない）★★
     `lib/grid-stats.js` の 帯の 平均 … ★丸めた 時だけ ≒ を 付ける★
       `(r===n ? '' : '≒') + fmt(r)`
     ⇒★丸めた 事が お客さんに 見える＝「黙って」では ない★＝D群
     ⇒★マスの 値では なく 帯（画面の下）の 字★ */
const 四桁の免除 = [
  { 道: 'lib/grid-stats.js', 訳: '帯の 平均 … ★丸めた 時だけ ≒ を 付ける（黙って 丸めない）★／マスの 値では ない' },
];
T('★★黙って 4桁に する 丸めが 本番の コードに 1つも 無い（免除は 理由つき 1か所）★★', () => {
  const なし = 四桁の免除.filter((x) => !x.訳 || !x.訳.trim());
  if (なし.length) throw new Error('★免除に 理由が 書いて いない★');
  const 免除の道 = new Set(四桁の免除.map((x) => x.道));
  const 全部 = 探す(素で読む).filter((x) => /\*\s*10000\s*\)\s*\/\s*10000/.test(x.字));
  const 今 = 全部.filter((x) => !免除の道.has(x.道));
  console.log('      4桁の 丸め … 全 ' + 全部.length + 'か所 ／ 免除 ' + (全部.length - 今.length)
    + 'か所（理由つき）／ ★黙って 丸めて いる ' + 今.length + 'か所★');
  if (今.length) {
    throw new Error('★黙って 4桁に する 丸めが ' + 今.length + 'か所 在る★\n'
      + 今.map((x) => '      ' + x.道 + ':' + x.行 + '  ' + x.字).join('\n')
      + '\n      ⇒★お客さんの 答えを 黙って 4桁に する 形は 使いません★');
  }
});

/* ══ ★自己試験＝★わざと 足して 赤に なる所まで 見る★ ══════════ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  const 差し替え = (道, 足す) => (p) => {
    const s = 素で読む(p);
    if (s === null) return null;
    return path.relative(ROOT, p).replace(/\\/g, '/') === 道 ? s + '\n' + 足す : s;
  };

  T('★★わざと 1か所 足したら 赤に なる（★一番 大事★）★★', () => {
    const 今 = 探す(差し替え('lib/ink.js', 'var x = Math.round(y * 100000) / 100000;'));
    const 新しい = 今.filter((x) => !一覧の鍵.has(鍵(x)));
    console.log('      … 足した 1か所が ★' + 新しい.length + '個★ 新しいと 出た');
    if (新しい.length !== 1) throw new Error('★' + 新しい.length + '個★（1個のはず）＝この 見張りは 何も 見て いない');
  });

  T('★7通り どの 書き方でも 見つかる★', () => {
    const 見本 = [
      ['①', 'var a = Math.round(z * 1000) / 1000;'],
      ['②', 'var b = Math.round(z * 10) / 10;'],
      ['③', 'var c = z.toFixed(3);'],
      ['④', 'var d = z.toPrecision(4);'],
      ['⑤', 'var e = Math.floor(z * 1000) / 1000;'],
      ['⑥', 'var f = parseFloat(z.toFixed(3));'],
      ['⑦', 'var g = Number(z.toFixed(3));'],
    ];
    const 出ず = [];
    for (const [札, 字] of 見本) {
      const 今 = 探す(差し替え('lib/ink.js', 字));
      if (!今.filter((x) => !一覧の鍵.has(鍵(x))).length) 出ず.push(札 + ' ' + 字);
    }
    console.log('      7通り 全部 試した … 見つからなかった ' + 出ず.length + '通り');
    if (出ず.length) throw new Error('★' + 出ず.length + '通り 見つからない★\n      ' + 出ず.join('\n      '));
  });

  /* ★★2026-09-08 に 書き替えた★★
       前の 自己試験は「A群を 1つ 消したら 減ると 言うか」だった。
       ⇒ A群は 4つとも ★外し終えた★ので、消す 物が もう 無い＝★この 試験は 意味を 失った★
       ⇒★守る 物が 変わったので 自己試験も 書き替える★
         新しい 守る 物＝★黙って 4桁に する 丸めが 戻って きたら 赤★ */
  T('★★4桁の 丸めを 戻したら 赤に なる（★一番 大事★）★★', () => {
    const 戻す = (p) => {
      const s = 素で読む(p);
      if (s === null) return null;
      if (path.relative(ROOT, p).replace(/\\/g, '/') !== 'exally-formula.js') return s;
      /* ★★探す 字が 古く なって いました（2026-09-09 に 直した）★★
         前は `String(_jsXirr(vals,dates))` を 探して いました。
         ⇒ XIRR に 門を 付けた 時 呼び方が `String(_jsXirr(組.値,組.日,見当))` に なり
           ★字が 当たらなく なって 自己試験が「戻したのに 0か所」と 言った★
         ⇒★本体は 緑・自己試験だけ 赤★＝★見張りが 見て いない事を 自己試験が 捕まえた★
         ⇒★字を 1つに 決め打ちせず、★在る 物を 見つけてから 包む★★ */
      const m = /String\(_jsXirr\(([^)]*)\)\)/.exec(s);
      if (!m) return s;                       /* ★見つからなければ そのまま★（下で 赤に なる） */
      return s.replace(m[0], 'String(Math.round(_jsXirr(' + m[1] + ')*10000)/10000)');
    };
    const 免除の道 = new Set(四桁の免除.map((x) => x.道));
    const 前 = 探す(素で読む).filter((x) => /\*\s*10000\s*\)\s*\/\s*10000/.test(x.字)).filter((x) => !免除の道.has(x.道));
    const 後 = 探す(戻す).filter((x) => /\*\s*10000\s*\)\s*\/\s*10000/.test(x.字)).filter((x) => !免除の道.has(x.道));
    if (前.length !== 0) throw new Error('★今 既に ' + 前.length + 'か所 在る★（0のはず）');
    if (後.length !== 1) throw new Error('★戻したのに ' + 後.length + 'か所★（1のはず）＝★見張りが 見て いない★');
    console.log('      … 黙って 丸める 所 0 → ' + 後.length + '（★戻したら 赤に なると 分かる★）');
  });

  /* ★★入れ子 2段を 読めるか（★1回目・2回目の 見落としと 同じ 形★）★★
       `Math.round(f(g(x))*10000)/10000` は ★正規表現では 拾えなかった★
       ⇒ 括弧を 数えて 読む 形に 直した＝★段数に 上限が 無い★事を ここで 確かめる */
  T('★★入れ子が 何段でも 拾える（2段・3段を わざと 作って 試す）★★', () => {
    const 足す = (p) => {
      const s = 素で読む(p);
      if (s === null) return null;
      if (path.relative(ROOT, p).replace(/\\/g, '/') !== 'exally-formula.js') return s;
      return 'var _2=Math.round(f(g(x))*10000)/10000;\n'
        + 'var _3=Math.round(f(g(h(x)))*10000)/10000;\n'
        + 'var _4=Math.round(f(g(h(i(x))))*10000)/10000;\n' + s;
    };
    const 免除の道 = new Set(四桁の免除.map((x) => x.道));
    const 後 = 探す(足す).filter((x) => /\*\s*10000\s*\)\s*\/\s*10000/.test(x.字)).filter((x) => !免除の道.has(x.道));
    if (後.length !== 3) throw new Error('★2段・3段・4段の うち 拾えたのは ' + 後.length + '個★（3個のはず）'
      + '\n      ⇒★入れ子の 段数で 見落とす＝1回目(XIRR)・2回目(BINOM)と 同じ 穴★');
    console.log('      … 2段/3段/4段 とも 拾えた（' + 後.length + '個）');
  });

  T('★★借り物（min.js）を 見て いない★★', () => {
    const 見た = 見るファイル().map((p) => path.relative(ROOT, p).replace(/\\/g, '/'));
    const 借り = 見た.filter((p) => p.endsWith('.min.js'));
    if (借り.length) throw new Error('★借り物を ' + 借り.length + '本 見て いる★ … ' + 借り.join(' '));
    if (fs.existsSync(path.join(ROOT, 'hyperformula.full.min.js'))
      && 見た.indexOf('hyperformula.full.min.js') >= 0) throw new Error('★借り物の エンジンを 見て いる★');
    console.log('      … ' + 見た.length + '本 中 借り物 0本');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
