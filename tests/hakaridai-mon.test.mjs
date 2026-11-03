/* hakaridai-mon.test.mjs — ★測り台が 本番と 違うのを 止める★（2026-09-10）
 *
 *  ★★2026-09-09〜10 の 1日で ★5回★ 転びました★★
 *    ①プラグインを ★3本しか★ 積んで いなかった（本番は 8本）… LINEST の 道具
 *    ②プラグインを ★1本も★ 積んで いなかった ………………… 丸めの 押し比べ（★#53 の 根拠★）
 *    ③`smartRounding` が ★既定（true）★（本番は false）…… 803.6538461538445 → ★803.65384615★
 *    ④★板ごと `setSheetContent` で 入れて いた★（本番は 1マスずつ）
 *       ⇒ 裸の `=LINEST(…)` が ★#VALUE!★ に なり
 *         ★「本番が 壊れて いる」と 報告する 一歩 手前★まで 行った
 *       ⇒★実配信を ブラウザで 押したら 4本とも 動いて いた★＝★私の 台の 産物★
 *    ⑤★`getCell` は 式の 答えを 持って いない★／`setCell(r,c,v)` は ★シート番号を 取らない★
 *       ⇒ 材料を 5マス ずらして 敷き ★「1マス目も 空＝穴だ」と 読みかけた★
 *
 *  ★★だから この 見張りが 在ります★★
 *    ★node の 台で 出した 数は ★画面の 数では ありません★★
 *    ⇒ 台の 作り方が 本番と 違えば ★止める★
 *    ⇒ 画面の 事を 言いたい なら ★ブラウザで 押す★（★その 断りを 道具に 書かせる★）
 *
 *  ★★何を 守るか（1つずつ 名指し）★★
 *    ①★本番が 読む プラグインの 数と 同じだけ 積んで いる★
 *    ②★`smartRounding:false` と `useArrayArithmetic:true` を 建て方に 書いて いる★
 *    ③★`setSheetContent` で 板ごと 入れる 道具は ★断りを 書く★★
 *    ④★「画面」「本番の 画面」と 名乗る 道具は ★node では ない★事を 確かめる★
 *
 *  ★★見て いない 範囲（★書かない 見張りは「全部 守った」と 読まれる★）★★
 *    ・★`docs/measured/osu-*.mjs` だけ★ 見ます（`toru-*.ps1` は 実Excel＝別の 話）
 *    ・★中の 計算が 合って いるかは 見て いません★＝★台の 作り方だけ★
 *    ・★ブラウザで 押したかは 機械では 分かりません★＝★断りが 在るかだけ★
 *
 *  使い方: node tests/hakaridai-mon.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ★本番が 読む プラグインの 数★（book.html から 数える＝★書き込まない★） */
function 本番のプラグイン数(html) {
  return new Set([...html.matchAll(/lib\/(formula-[a-z]+)-plug\.js/g)].map((m) => m[1])).size;
}
/* ★本番の 建て方★ */
function 本番のbuildEmpty(html) {
  const m = /buildEmpty\(\{([\s\S]{0,300}?)\}\)/.exec(html);
  return m ? m[1].replace(/\s+/g, '') : null;
}

const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf-8');
const 要る本数 = 本番のプラグイン数(book);
const 本番の建て方 = 本番のbuildEmpty(book);

/* ★見る 道具★＝`docs/measured/osu-*.mjs` */
function 道具ら() {
  const d = path.join(ROOT, 'docs/measured');
  if (!fs.existsSync(d)) return [];
  return fs.readdirSync(d).filter((f) => /^osu-.*\.mjs$/.test(f))
    .map((f) => ({ 名: f, 道: path.join(d, f), 字: fs.readFileSync(path.join(d, f), 'utf-8') }));
}

/* ══ ★免除（★理由つきで 名指し★／黙って 見逃さない）★ ══════════ */
const 免除 = [
  /* ★★osu-wakeru.mjs は 2026-09-15 に 免除から 外しました★★
     ★前の 訳★ …「見るのは 返って きた 物の 種類だけなので
             プラグインの 数も smartRounding も 答えを 変えません」
     ★★これは 間違って いました★★（実測・2026-09-15）
       本番の 道に 乗せたら ★「誤りしか 返らない」が 71個 → 68個★
       変わった 3個 … ★ACOTH / ATANH / F.INV★
       ＝★軽い 建て方だと「誤り」に 見えて いた★
       ＝★種類だけ 見る 道具でも 建て方で 種類が 変わる★
     ⇒★★「答えを 使わないから 建て方は 適当で よい」は 成り立ちません★★
     ⇒★`docs/measured/honban-no-michi.mjs` に 乗せて 免除を 1つ 減らしました★ */
  { 名: 'osu-oufuku.mjs',
    訳: '★2026-09-10 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★うちの 書き出しの 道（GridXlsx → XlsxIO）で xlsx を 作る★のと '
      + '★実Excel が 開いて 測った 紙と 突き合わせる★の 2つだけ。'
      + '⇒ 式を 1つも 計算しないので ★プラグインを 積んでも 答えは 1文字も 変わりません★。'
      + '⇒★答えの 正しさは 実Excel が 出して います★（docs/measured/toru-oufuku.ps1 で 開いて 測った）。'
      + '★これこそが この 道具の 役目★＝★うちの 台では なく 実Excel に 答えさせる★' },
  { 名: 'osu-shisuu-mitame.mjs',
    訳: '★2026-09-10 の 道具★＝★エンジンを 1度も 呼びません★。'
      + '押すのは `book.html` の `forDisplay`＝★エンジンの 答えが 出た 後、画面に 字を 作る 段★だけ。'
      + '⇒ 数は ★実Excel に 打たせた 紙★から 読み、うちは ★字を 作る 段に 直に 渡す★（本番の 画面と 同じ 渡り方）。'
      + '⇒★エンジンを 積んでも この 道具の 答えは 1文字も 変わりません★（通らない 段だから）。'
      + '★そして 同じ 10本を ★ブラウザで 打って 絵で 確かめて あります★'
      + '（`docs/measured/e-shisuu-mitame-2026-09-10.png`）＝★台の 数だけで 済ませて いません★' },
  { 名: 'osu-afurenaoshi.mjs',
    訳: '★2026-09-11 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★本物の ブラウザで book.html を 開き、マスを 押して 字を 打つ★だけ。'
      + '⇒ 計算するのは ★画面の 中の エンジン（本番と 同じ 積み方）★＝'
      + '★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '⇒★だから プラグインの 数も smartRounding も この 道具の 答えを 1文字も 変えません★。'
      + '★この 道具の 数は 画面の 数です★＝台の 数より 強い 物で、'
      + '★絵も 出します★（docs/measured/e-afurenaoshi-2026-09-11.png）' },
  { 名: 'osu-excel-kara.mjs',
    訳: '★2026-09-11 の 道具★＝★getCellValue は 呼びます★。'
      + '★でも 自分の 台では ありません★＝`page.evaluate` の 中で '
      + '★画面（book.html）が 建てた エンジン★に 聞いて います。'
      + '⇒ その エンジンは ★本番そのもの★＝プラグイン 8本・smartRounding:false も 本番の まま。'
      + '⇒★この 道具は buildEmpty を 1度も 呼びません★（自分では 建てない）。'
      + '⇒★だから「本番と 違う 台で 押す」事は 起き得ません★。'
      + '★2026-09-11 の 書き直し★ … 最初は ★呼ばないと 書いて いました★が、'
      + '★門に 嘘だと 見抜かれました★（getCellValue が 字の 上に 在る）。'
      + '★実物に 合わせて 書き直しました★＝★言い切りの 字を そのまま 引かない★'
      + '（引くと 門が それを 言い切りと 読みます）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-excel-kara.ps1 が 作って 測った 紙と 突き合わせる）' },
  { 名: 'osu-moji-no-kotae.mjs',
    訳: '★2026-09-11 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★本物の ブラウザで book.html を 開き、マスを 押して 字を 打ち★、'
      + '★画面が 本当に 描いた 字（fillText）★を 取る事だけ。'
      + '⇒ 計算するのは ★画面の 中の エンジン＝本番そのもの★'
      + '（プラグイン 8本・smartRounding:false も 本番の まま）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-moji-no-kotae.ps1 で 14本 打たせて 測った）' },
  { 名: 'osu-jitsubutsu-deruji.mjs',
    訳: '★2026-09-11 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★本物の ブラウザで 司さんの 実物（の 写し）を お客さんの 入口から 開き★、'
      + '★画面が 本当に 描いた 字★を 取る事だけ。'
      + '⇒ 計算するのは ★画面の 中の エンジン＝本番そのもの★'
      + '（プラグイン 8本・smartRounding:false も 本番の まま）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-jitsubutsu-deruji.ps1 が 同じ ファイルを 開いて 測った）。'
      + '★司さんの 実物と その 紙は repo に 入れて いません★（取引先の 名前・金額）。' },
  { 名: 'osu-shikaku-kansuu.mjs',
    訳: '★2026-09-13 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★実Excel に 打って 取った 紙（golden-shikaku-kansuu-2026-09-13.tsv）★と '
      + '★見張り（tests/shiki-kansuu.test.mjs）の 中の 37通り★を ★字で 突き合わせる★ 事だけ。'
      + '⇒★エンジンは 借り物も 自前も 1つも 呼びません★（fs と path しか 使って いません）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-shikaku-kansuu.ps1 で 40通り 打たせて 測った）。' },
  { 名: 'osu-tsunagi.mjs',
    訳: '★2026-09-11 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★実Excel に 打って 取った 紙（golden-tsunagi-2026-09-11.tsv）★と '
      + '★見張り（tests/shiki-keisan.test.mjs）の 中の 162通り★を ★字で 突き合わせる★ 事だけ。'
      + '⇒★エンジンは 借り物も 自前も 1つも 呼びません★（fs と path しか 使って いません）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-tsunagi.ps1 で 162通り 打たせて 測った）。' },
  { 名: 'osu-sashikomi.mjs',
    訳: '★2026-09-13 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★実Excel に 打って 取った 紙（golden-sashikomi-2026-09-13.tsv）★と '
      + '★見張り（tests/shiki-sashikomi.test.mjs）の 中の 17通り★を ★字で 突き合わせる★ 事だけ。'
      + '⇒★エンジンは 借り物も 自前も 1つも 呼びません★（fs と path しか 使って いません）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-sashikomi.ps1 で 17通り 打たせて 測った）。' },
  { 名: 'osu-afure.mjs',
    訳: '★2026-09-13 の 道具★＝★buildEmpty を 1度も 呼びません★（自分の 台を 建てない）。'
      + 'やる事は ★実Excel に 打って 取った 紙（golden-afure-2026-09-13.tsv）★と '
      + '★自前の 溢れの 台（lib/shiki-afure.js）の 答え★を ★字で 突き合わせる★ 事だけ。'
      + '⇒★読み込むのは lib/shiki-keisan.js と lib/shiki-afure.js の 2本だけ★＝'
      + '★どちらも 借り物を 1行も 読みません★（借り物を 外す 為に 自分で 書いた 台）。'
      + '⇒★この 道具が 積む 物は 1つも 在りません★（積みようが ない）。'
      + '★答えの 正しさは 実Excel が 出して います★'
      + '（docs/measured/toru-afure.ps1 で 16通り 打たせて 測った）。' },
  /* ★★2026-09-15 に 足した 3本★★（★司さんの実物を 読む 道具★）
     ★3本とも エンジン（借り物）を 1度も 呼びません★＝★対象外★
     ＝`buildEmpty` も `getCellValue` も 1つも 書いて いません（★機械が 上で 確かめます★） */
  { 名: 'osu-jitsubutsu-mitame.mjs',
    訳: '★2026-09-15 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★SheetJS の `SSF.format`★ と ★自前の 書式の 台（lib/shoshiki.js）★を '
      + '★字で 突き合わせる★ 事だけ（司さんの実物 16,536マス）。'
      + '⇒★くらべる 相手は どちらも JS★＝★式を 1本も 計算しません★。'
      + '⇒★積む 物は 1つも 在りません★（積みようが ない）。' },
  { 名: 'osu-jitsubutsu-hyou-sansho.mjs',
    訳: '★2026-09-15 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★表の 参照（Table1[…]）が A1 に 直るか★を `lib/table-refs.js` に 直させて '
      + '★本数を 数える★ 事だけ。⇒★式を 1本も 計算しません★（形を 見るだけ）。'
      + '⇒★積む 物は 1つも 在りません★。' },
  { 名: 'osu-jitsubutsu-dodai.mjs',
    訳: '★2026-09-15 の 道具★＝★エンジンを 1度も 呼びません★。'
      + '★これは 道具では なく ★押し方そのもの★★＝'
      + '「表の 参照を 直す／板またぎに 染まった マスを 印す／表を 建てて 押す」を '
      + '★1つの 物に 出した 物★です（★掘る 道具が 書き起こすと 印が 抜ける★ので）。'
      + '⇒★呼ぶのは 自前の 土台（lib/shiki-hyou.js）だけ★＝★借り物は 建てません★。'
      + '⇒★積む 物は 1つも 在りません★。' },
  { 名: 'osu-jitsubutsu-awanai37.mjs',
    訳: '★2026-09-15 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★`osu-jitsubutsu-jizen-dake.mjs` で 合わなかった 分の 中を 見る★ 事＝'
      + '★自前の 土台（lib/shiki-hyou.js）だけ★で 押して ★関数の 名前と 本数を 数える★。'
      + '⇒★借り物を 建てないのが この 道具の 目的★です（★借り物と くらべて いません★）。'
      + '⇒★積む 物は 1つも 在りません★。' },
  { 名: 'osu-jitsubutsu-jizen-dake.mjs',
    訳: '★2026-09-15 の 道具★＝★エンジンを 1度も 呼びません★。'
      + 'やる事は ★自前の 土台（lib/shiki-hyou.js）だけで 実物の 式を 押す★ 事＝'
      + '★借り物を 建てないのが この 道具の 目的★です（★借り物と くらべて いません★）。'
      + '⇒ 正しさの 相手は ★ファイルに 焼かれた 実Excel の 答え★（`c.v`）。'
      + '⇒★積む 物は 1つも 在りません★。' },
];

console.log('\n[hakaridai-mon] ★測り台が 本番と 違うのを 止める★');
console.log('  ★本番★ … プラグイン ' + 要る本数 + '本 ／ buildEmpty ' + (本番の建て方 || '(読めない)'));

T('★道具を 1本でも 見つけて いる（★空振りして いない★）★', () => {
  const n = 道具ら().length;
  if (!n) throw new Error('★`docs/measured/osu-*.mjs` が 1本も 無い★');
  if (!要る本数) throw new Error('★book.html から プラグインの 数を 数えられない★');
  if (!本番の建て方) throw new Error('★book.html の buildEmpty が 読めない★');
  console.log('      … 道具 ' + n + '本');
});

T('★★本番の 建て方に smartRounding:false と useArrayArithmetic:true が 在る★★', () => {
  /* ★本番が 変わったら ここで 気づく★（★下の 試験の 前提★） */
  for (const 要 of ['smartRounding:false', 'useArrayArithmetic:true']) {
    if (本番の建て方.indexOf(要) < 0) {
      throw new Error('★本番の buildEmpty に ' + 要 + ' が 無い★＝★下の 試験の 前提が 崩れた★');
    }
  }
  console.log('      … 2つとも 在る');
});

/* ══ ★★本番の 道を 通って いるか★★ ══（2026-09-15 に 強めました）
     ★前は★ …「道具の 中に 書いて ある 建て方が 本番と 同じか」を 見て いました
       ＝★写しを 許して いた★。実際 ★同じ 建て方が 6本★ 在りました。
     ★今は★ …★★`docs/measured/honban-no-michi.mjs` を 通って いるか★★だけ 見ます
       ＝★自前で 建てて いる 道具が 1本でも 在れば 赤★
     ★なぜ 強めたか★
       ★悪い 書き方は 何通りでも 書ける／良い 呼び方は 1つ★
       （[[feedback_warui_kakikata_wo_sagasu_yori_yoi_yobikata_wo_kazoero]]）
     ★実際に 出た 害★（2026-09-15・実測）
       `osu-wakeru.mjs` は ★軽い 建て方★で 建てて いました
       ⇒★「誤りしか 返らない」が 71個 → 本番の 道で 68個★
       ⇒★ACOTH / ATANH / F.INV を 「誤り」と 数えて いた★ */
T('★★道具は 本番の 道（honban-no-michi）を 通って いる★★', () => {
  const 悪い = [];
  for (const t of 道具ら()) {
    if (免除.some((x) => x.名 === t.名)) continue;
    const 素 = 注記を外す(t.字);
    if (素.indexOf('honban-no-michi') >= 0) continue;          /* ★通って いる★ */
    const 自前 = [];
    if (/\.buildEmpty\s*\(/.test(素)) 自前.push('buildEmpty');
    if (/-plug\.js/.test(素)) 自前.push('plug を 自前で つなぐ');
    if (自前.length) 悪い.push(t.名 + '（' + 自前.join('／') + '）');
  }
  if (悪い.length) {
    throw new Error('★' + 悪い.length + '本が 自前で 建てて います★\n      '
      + 悪い.join('\n      ')
      + '\n      ⇒★`docs/measured/honban-no-michi.mjs` の `建てる()` を 使って ください★'
      + '\n      ⇒★違いが 要るなら ★別の 道では なく 口の 引数★ に して ください★'
      + '\n      ⇒ 2026-09-15 実測 … 軽い 建て方だと ACOTH/ATANH/F.INV を 「誤り」と 数えて いた');
  }
  console.log('      … 測り台 ' + 道具ら().filter((t) => !免除.some((x) => x.名 === t.名)).length
    + '本とも 本番の 道を 通って いる ／ 免除 ' + 免除.length + '本（理由つき）');
});

/* ★★本番の 道 そのものが 本番と 同じか★★（★台は 1本★なので ここだけ 見る） */
T('★★本番の 道の 建て方が 本番と 同じ（smartRounding:false）★★', () => {
  const 道 = 注記を外す(fs.readFileSync(path.join(ROOT, 'docs/measured/honban-no-michi.mjs'), 'utf-8'));
  const 建て = (/buildEmpty\(\{([\s\S]{0,300}?)\}\)/.exec(道) || [])[1];
  if (!建て) throw new Error('★本番の 道に buildEmpty が 無い★');
  const x = 建て.replace(/\s+/g, '');
  for (const 要 of ['smartRounding:false', 'useArrayArithmetic:true']) {
    if (x.indexOf(要) < 0) throw new Error('★本番の 道に ' + 要 + ' が 無い★'
      + '＝★803.6538461538445 が 803.65384615 に なります（2026-09-09 実測）★');
  }
  /* ★本番の 道が プラグインの 本数を ★書き込んで いない★ 事★
     ＝★book.html から 数えて いる★（本番が 増えたら 自分で 止まる） */
  if (道.indexOf('本番のプラグイン数') < 0) {
    throw new Error('★本番の 道が 本数を book.html から 数えて いない★');
  }
  console.log('      … 本番の 道 1本を 見た（smartRounding:false／useArrayArithmetic:true／本数は book.html から）');
});

T('★★`setSheetContent` で 板ごと 入れる 道具は 断りを 書いて いる★★', () => {
  /* ★本番は `setCell`／`setCellFormula` で ★1マスずつ★ です★
     ⇒ 板ごと 入れると ★式が 表の 中を 指したり 溢れの 判じが 変わったり★します
     ⇒★使うなと は 言いません★（速いので）。★「画面の 数では ない」と 書かせます★ */
  const 悪い = [];
  for (const t of 道具ら()) {
    if (免除.some((x) => x.名 === t.名)) continue;
    if (t.字.indexOf('setSheetContent') < 0) continue;
    const 断り = /画面の 数では ない|画面の 答えでは ない|ブラウザで 押/.test(t.字);
    if (!断り) 悪い.push(t.名);
  }
  if (悪い.length) {
    throw new Error('★' + 悪い.length + '本が 板ごと 入れて いるのに 断りが 無い★\n      '
      + 悪い.join('\n      ')
      + '\n      ⇒★「★この 台の 数は 画面の 数では ありません★」と 書いて ください★'
      + '\n      ⇒ 2026-09-10 … 板ごと 入れて 裸の =LINEST が #VALUE! に なり'
      + '\n         ★本番が 壊れて いると 報告する 一歩 手前★まで 行きました');
  }
  console.log('      … 板ごと 入れる 道具は 全部 断りつき');
});

/* ★★「免除」と「対象外」を 分けます（2026-09-10）★★
   ★この 門が 見るのは ★エンジンを 建てて 式を 押す 道具★です。
   ★エンジンを 1度も 呼ばない 道具★は ★そもそも 測り台では ありません★＝★対象外★。
   ⇒ 前は それも「免除」に 入れて いたので ★免除が 3本に 増え、
     「免除に 逃げて いないか」の 見張りが 赤に なりました★（★門が 正しい★）
   ⇒★言い張りに しない★＝★機械で 判じます★
     `buildEmpty` も `getCellValue` も ★1つも 書いて いない★＝エンジンを 触らない
   ⇒★逃げ道に しない★＝理由も 要るし、★本当に 触って いない事を 数える★ */
/* ★改行を 逃がし（バックスラッシュ n）で 書きません★
   ＝2026-09-07/10 に ★heredoc で 生の 改行に 化ける★のを ★4回★ 踏んだ */
const 改行と字下げ = String.fromCharCode(10) + '      ';

function エンジンを触るか(t) {
  return t.字.indexOf('buildEmpty') >= 0 || t.字.indexOf('getCellValue') >= 0;
}

T('★★「対象外」は 本当に エンジンを 触って いない（★言い張りに しない★）★★', () => {
  const 嘘 = [];
  for (const x of 免除) {
    const t = 道具ら().find((y) => y.名 === x.名);
    if (!t) continue;
    /* ★★探す 字を 決めた 時点で 答えが 決まる★★（2026-09-10 に 踏んだ）
       前は「種類だけ」も 拾って いて、★エンジンを 触る 道具まで 対象外と 読んで★
       ★中身は 正しいのに 赤★に なりました。
       ⇒★見るのは 1つの 言い切りだけ★＝「エンジンを 1度も 呼びません」
       ⇒★それ以外の 免除は「エンジンを 触る けれど 免除」として 数える★（上の 見張りが 数を 見る） */
    if (x.訳.indexOf('エンジンを 1度も 呼びません') < 0) continue;
    if (エンジンを触るか(t)) 嘘.push(x.名 + ' … ★「エンジンを 呼ばない」と 書いて あるのに buildEmpty か getCellValue が 在る★');
  }
  if (嘘.length) throw new Error('★' + 嘘.length + '本★' + 改行と字下げ + 嘘.join(改行と字下げ));
  console.log('      … 書いて ある 通り 触って いない');
});

T('★免除は 全部 理由つき／★免除に 逃げて いない★★', () => {
  const なし = 免除.filter((x) => !x.訳 || !x.訳.trim());
  if (なし.length) throw new Error('★' + なし.length + '本に 理由が 無い★');
  /* ★エンジンを 触る 道具だけで 数える★
     ＝★触らない 道具は そもそも 測り台では ない★（対象外）ので 分母にも 分子にも 入れない */
  const 測り台 = 道具ら().filter(エンジンを触るか);
  const 本当の免除 = 免除.filter((x) => {
    const t = 道具ら().find((y) => y.名 === x.名);
    return t && エンジンを触るか(t);
  });
  const 対象外 = 免除.length - 本当の免除.length;
  if (本当の免除.length > Math.max(1, Math.floor(測り台.length / 3))) {
    throw new Error('★免除が 多すぎる（' + 本当の免除.length + ' ／ 測り台 ' + 測り台.length + '）★'
      + 改行と字下げ + '⇒★エンジンを 触る 道具を 免除するのは よほどの 時だけ★');
  }
  console.log('      … 測り台 ' + 測り台.length + '本 ／ ★免除 ' + 本当の免除.length + '本★ ／ 対象外 ' + 対象外 + '本（エンジンを 触らない）');
});

/* ══ ★★⑤誤りの マスを `v` で 比べて いないか★★ ══（2026-09-15 に 足した）
     ★★何が 起きたか★★
       ファイルの 誤りの マスは ★`t:'e'` ／ `v:23`（数）／ `w:"#REF!"`（字）★
       （実測＝司さんの 実物に 69マス）。
       ⇒★`v` の まま 突き合わせると `#REF!` と 23 を 比べる★
       ⇒★★うちが 正しいのに「合わない」と 出る★★＝★★偽の 負け★★
       ★2026-09-15 実測★＝`=SUBTOTAL(109,#REF!)` 1本が これで 落ちて いた
         （★うちの 答えも 実Excel も #REF!★＝★道具の せいだった★）
     ★★決め★★ … ★`値たち[…].v` を 直に 読まない★＝`土台.正の値(セル)` を 通す
     ★見て いない 事★ … ★他の 読み方（`セル.v` を 別の 名前で 受ける 等）は 見て いません★
       ⇒★字で 探して いる 限り 抜けます★＝★本当の 守りは「合った 本数」★ */
T('★★誤りの マスを `v` で 直に 比べて いない★★（★偽の 負けを 止める★）', () => {
  const 悪 = /値たち\s*\[[^\]]*\]\s*\.\s*v\b/;
  const 当 = [];
  for (const t of 道具ら()) {
    const 素 = 注記を外す(t.字);   /* ★外す 道具は 共通部品★（自前で 書かない） */
    if (悪.test(素)) 当.push(t.名);
  }
  console.log('      … 見た 道具 ' + 道具ら().length + '本／★`値たち[…].v` を 直に 読む 道具 '
    + 当.length + '本★');
  if (当.length) {
    throw new Error('★誤りの マスは `v` が 数（#REF! は 23）★＝★`土台.正の値(セル)` を 通して ください★'
      + '\n        ' + 当.join('\n        '));
  }
});

/* ══ ★★⑥`!` を 板の 印と 読む 所で `#REF!` を 外して いるか★★ ══（2026-09-15 に 足した）
     ★★何が 起きたか★★
       `板か()` が `/[A-Za-z0-9_…']+!/` で 見て いた
       ⇒★`#REF!` の `REF!` を「板の 名前＋！」と 読んで いた★
       ⇒★実物 93本を「板を またぐ」に 入れて いた★＝★押せるのに 押して いなかった★
         （＋★それに 頼った 式まで 染めて いた★＝染まり 315 → 312）
     ★★しかも 直した 後も 2本目が 残って いました★★
       `osu-jitsubutsu-hyou-sansho.mjs` に ★同じ 形の 写し★が 在った
       ＝★作る道が 2本★＝★片方だけ 直すと 黙って 食い違う★
     ★★決め★★ … ★`…+!/` の 形で 板名を 見る 所は `#REF!` を 先に 外す★
     ★見て いない 事★
       ・★`#NAME?` `#N/A` は `!` を 持たない★ので 当たりません（★`#REF!` だけが `!` を 持つ★）
       ・★`shiki-sansho` は `#REF!A1` を 板「#REF」と 読みます★
         ＝★Excel の 板消しの 形★＝★今の 実物には 0本★（実測）
         ⇒★㋖（板またぎ）で 板を 引く 時に「無い 板＝#REF!」を 出す事で 受けます★ */
T('★★`!` を 板の 印と 読む 所は `#REF!` を 外して いる★★（★作る道を 2本に しない★）', () => {
  const d = path.join(ROOT, 'docs/measured');
  const 見た = [], 悪 = [];
  for (const f of fs.readdirSync(d)) {
    if (!/\.(mjs|js)$/.test(f)) continue;
    const 素 = 注記を外す(fs.readFileSync(path.join(d, f), 'utf-8'));
    /* ★`…']+!/` の ように「何かの 字＋！」で 板名を 見て いる 行★ */
    for (const 行 of 素.split('\n')) {
      if (!/\+!\//.test(行)) continue;
      見た.push(f);
      if (行.indexOf('#REF!') < 0) 悪.push(f + '  ' + 行.trim().slice(0, 72));
    }
  }
  console.log('      … ★板名を `+!` で 見て いる 行 ' + 見た.length + '本／'
    + '★`#REF!` を 外して いない ' + 悪.length + '本★');
  if (悪.length) {
    throw new Error('★`#REF!` の `REF!` を 板名と 読みます★＝★`.replace(/#REF!/g, \'\')` を 通して ください★'
      + '\n        ' + 悪.join('\n        '));
  }
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★smartRounding を 抜いたら 赤に なる★★', () => {
    const 写し = 道具ら().filter((t) => !免除.some((x) => x.名 === t.名))[0];
    if (!写し) throw new Error('★見る 道具が 無い★');
    const 抜いた = 写し.字.replace(/smartRounding:\s*false,?/g, '');
    const 建て = (/buildEmpty\(\{([\s\S]{0,300}?)\}\)/.exec(抜いた) || [])[1] || '';
    if (建て.replace(/\s+/g, '').indexOf('smartRounding:false') >= 0) {
      throw new Error('★抜いたのに まだ 在る＝この 試験は 空振り★');
    }
    console.log('      … ' + 写し.名 + ' から 抜くと 見つからなく なる（★赤に なる形★）');
  });

  T('★★断りを 消したら 赤に なる★★', () => {
    const 板 = 道具ら().filter((t) => t.字.indexOf('setSheetContent') >= 0
      && !免除.some((x) => x.名 === t.名));
    if (!板.length) { console.log('      … 板ごと 入れる 道具が 無い（この 試験は 当たらない）'); return; }
    const 消した = 板[0].字.replace(/画面の 数では ない|画面の 答えでは ない|ブラウザで 押/g, '');
    if (/画面の 数では ない|画面の 答えでは ない|ブラウザで 押/.test(消した)) {
      throw new Error('★消したのに まだ 在る＝この 試験は 空振り★');
    }
    console.log('      … ' + 板[0].名 + ' から 断りを 消すと 見つからなく なる');
  });

  T('★★`値たち[…].v` を 書き戻したら 赤に なる★★', () => {
    const 悪 = /値たち\s*\[[^\]]*\]\s*\.\s*v\b/;
    const 写し = '    const 正 = 値たち[a] ? 値たち[a].v : undefined;';
    if (!悪.test(写し)) throw new Error('★探す 形が 当たって いない＝この 試験は 空振り★');
    const 直した = '    const 正 = 土台.正の値(値たち[a]);';
    if (悪.test(直した)) throw new Error('★直した 形まで 赤に する＝要らない 赤★');
    console.log('      … 直書きは 捕まえ／`正の値()` は 通す');
  });

  T('★★`#REF!` を 外すのを やめたら 赤に なる★★', () => {
    const 悪 = (行) => /\+!\//.test(行) && 行.indexOf('#REF!') < 0;
    if (!悪("const 板か = (f) => /[A-Za-z0-9_']+!/.test(裸に(f));")) {
      throw new Error('★探す 形が 当たって いない＝この 試験は 空振り★');
    }
    if (悪("const 板か = (f) => /[A-Za-z0-9_']+!/.test(裸に(f).replace(/#REF!/g, ''));")) {
      throw new Error('★直した 形まで 赤に する＝要らない 赤★');
    }
    console.log('      … 外して いない 形は 捕まえ／外した 形は 通す');
  });

  T('★本番の 数を 書き込んで いない（★book.html から 数えて いる★）★', () => {
    const 字 = fs.readFileSync(path.join(ここ, 'hakaridai-mon.test.mjs'), 'utf-8');
    const 本体 = 字.slice(字.indexOf('function 本番のプラグイン数'));
    if (/=== *8\b|!== *8\b|要る本数 *= *8/.test(本体)) {
      throw new Error('★8 を 書き込んで いる★＝★本番が 増えたら 追えない★');
    }
    console.log('      … 数は book.html から 数えて いる（今 ' + 要る本数 + '本）');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
