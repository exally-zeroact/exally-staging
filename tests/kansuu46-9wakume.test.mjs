/* kansuu46-9wakume.test.mjs — ★9枠目の 紙で 押す★（2026-09-18）
 *
 *  ★★なぜ この 試験が 紙と ★同じ commit★ で 出るか★★
 *    2026-09-18、`tests/` の どれからも 名前で 読まれない 紙が ★13枚 810行★
 *    在ると 分かりました。★その うち 3枚は その日 取った 物★です。
 *    ⇒★★「紙を 取る」で 終わりに して いた★★
 *    ⇒★決め★ ... ★紙を 取ったら ★同じ commit で★ それを 名前で 読む 試験も 出す★
 *
 *  ★★読む 紙★★ `docs/measured/golden-kansuu-9kaime-2026-09-18.tsv`
 *    ＝★実Excel 16.0 build 20326 ／ powershell.exe 5.1 ／ ANSI 932★
 *
 *  ★★材料（★紙の 頭に 書いて あります★＝道具が 決めて いません）★★
 *    A1:A5 = 1,2,3,4,5 ／ B1:B5 = 1,2,(=1/0),4,5
 *    C1:C5 = 1,3,5,7,9 ／ D1:D5 = 9,7,5,3,1
 *    ★F1:G2 = 1,2／10,20★（★行ごと 3・30 ／ 列ごと 11・22＝分かれる★）
 *    ★式は J1★（★材料の 外★）
 *
 *  ★★押す 行の 選び方★★
 *    ★台が 外側の 関数を 知って いる 行だけ★（★知らない 物は 押しません★）
 *    ⇒★「押して いない 行」も 数で 出します★（★黙って 飛ばさない★）
 *
 *  ★★合わせ方★★
 *    ・数どうしは ★数★で（紙は PowerShell の 'R'・台は JS の 一番 短い 書き方）
 *    ・誤りは ★「出る字」★と（答えの 欄は -2146826281 の ような 番号）
 *    ・★「打てません」と 出た 行は 押しません★（★実Excel が 式として 受けない★）
 *
 *  ★★まだ 合わない 物は 名指しで 許します★★
 *    ＝★数で 許すと 別の 物が 壊れても 気づけません★
 *    ＝★合ったら 許しを 外させる 門★も 付けます
 *
 *  使い方: node tests/kansuu46-9wakume.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));
const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));
const Tsu = require_(path.join(ROOT, 'lib/shiki-tsunagi.js'));

let pass = 0, fail = 0;
const T = (n, よい, 添え) => {
  if (よい) { pass++; console.log('  ok   ' + n); }
  else { fail++; console.log('  NG   ' + n + (添え ? '\n       ' + 添え : '')); }
};

console.log('');
console.log('[kansuu46-9wakume] ★9枠目の 紙で 押す★');

const 紙道 = path.join(ROOT, 'docs/measured/golden-kansuu-9kaime-2026-09-18.tsv');
T('★紙が 在る★', fs.existsSync(紙道), 紙道);
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/);
const 柱 = (行.find((l) => l.startsWith('# 番')) || '').replace(/^#\s*/, '').split('\t');
const 式列 = 柱.indexOf('式'), 答列 = 柱.indexOf('答え'), 字列 = 柱.indexOf('出る字');
T('★柱が 読めた★（式／答え／出る字）', 式列 >= 0 && 答列 >= 0 && 字列 >= 0, 柱.join(' / '));

/* ★★紙の 頭に 版が 書いて あるか★★（★版が 無い 紙は 突き合わせに 使えません★） */
const 頭 = 行.filter((l) => l.startsWith('#')).join('\n');
T('★紙の 頭に ★Excel の 版★が 在る★', 頭.indexOf('どの Excel か') >= 0);
T('★紙の 頭に ★貝殻の 版★が 在る★', 頭.indexOf('どの 貝殻か') >= 0);
T('★紙の 頭に ★文字コード★が 在る★', 頭.indexOf('どの 文字コードか') >= 0);

const 組 = [];
for (const l of 行) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  const 式 = c[式列];
  if (!式 || !式.startsWith('=')) continue;
  組.push({ 番: c[0], 式: 式, 答: c[答列], 字: c[字列] });
}
console.log('  ★紙の 行 ... ' + 組.length + '行★');
const 行の本数 = 61;
T('★★紙が ' + 行の本数 + '行★★（★増えても 減っても 赤★）',
  組.length === 行の本数, '出た ' + 組.length + '行');

/* ★台が 知る 名前★（土台 ＋ 皮 ＋ ★板の 特別な 形★） */
const 台 = new Set([
  ...Object.keys(K.表 || {}),
  ...Tsu.名前たち(),
  ...(H.特別な形 || []),
].map((x) => String(x).toUpperCase()));
T('★台が 300個 以上 読めた★', 台.size >= 300, '台 ' + 台.size + '個');

/* ══ ★まだ 合わない 物（★名指しで 許す★）★ ══ */
/* ══ ★★許しは 0本に なりました★★（2026-09-18）
     ★最後の 1件★ `=LAMBDA(x,x*2)(3)` ... ★読む 側を 直して 合いました★
     ⇒★★繋ぐ 前の 関門が 無くなりました★★
     ⇒★「合ったのに まだ の まま」で 機械が 教えました★ */
const まだ = new Map([
  /* ══ ★★ラムダの 一族の 許しは 外しました★★（2026-09-18）
       ＝★9本 とも 直りました★（MAP／REDUCE／SCAN／BYROW／BYCOL／MAKEARRAY／LAMBDA）
       ＝★★「合ったのに まだ の まま」で 機械が 教えました★★
       ⇒★許しを 増やす 時だけ 赤に する 門は 甘い★＝★減った 時も 赤★ */
]);

/* ══ ★材料（★紙の 頭から★） ══ */
function 板を作る() {
  const h = H.表();
  for (let i = 0; i < 5; i++) h.打つ('A' + (i + 1), String(i + 1));
  h.打つ('B1', '1'); h.打つ('B2', '2'); h.打つ('B3', '=1/0');
  h.打つ('B4', '4'); h.打つ('B5', '5');
  const C = [1, 3, 5, 7, 9], D = [9, 7, 5, 3, 1];
  for (let i = 0; i < 5; i++) { h.打つ('C' + (i + 1), String(C[i])); h.打つ('D' + (i + 1), String(D[i])); }
  h.打つ('F1', '1'); h.打つ('G1', '2'); h.打つ('F2', '10'); h.打つ('G2', '20');
  return h;
}
{
  const h = 板を作る();
  h.打つ('J1', '=SUM(A1:A5)');
  T('★材料が 入った★（=SUM(A1:A5) が 15）', String(h.字('J1')) === '15', '出た「' + h.字('J1') + '」');
  const g = 板を作る();
  g.打つ('J1', '=SUM(F1:G2)');
  T('★足した 材料も 入った★（=SUM(F1:G2) が 33）', String(g.字('J1')) === '33', '出た「' + g.字('J1') + '」');
}

/* ══ ★1行ずつ 押す★ ══ */
let 合 = 0, 違 = 0, 押さない = 0, 打てない = 0;
const 外れ = [], 許した = [], 直った = [];
for (const q of 組) {
  if (String(q.答).indexOf('打てません') >= 0) { 打てない++; continue; }
  const m = /^=([A-Z][A-Z0-9._]*)\(/.exec(q.式);
  const 外側 = m ? m[1] : null;
  if (!外側 || !台.has(外側)) { 押さない++; continue; }
  const h = 板を作る();
  h.打つ('J1', q.式);
  let 生 = null;
  try { 生 = h.値 ? h.値('J1') : null; } catch (e) { /* 字で 見る */ }
  const 字 = String(h.字('J1'));
  const 台の数 = (生 && 生.型 === '数') ? 生.値 : null;
  const 紙の数 = (q.答 !== '' && isFinite(Number(q.答)) && String(q.字).charAt(0) !== '#') ? Number(q.答) : null;
  let よい;
  if (台の数 !== null && 紙の数 !== null) {
    よい = (台の数 === 紙の数)
      || Math.abs(台の数 - 紙の数) <= Math.max(1e-9, Math.abs(紙の数) * 1e-9);
  } else {
    よい = (字 === String(q.字));
  }
  const 訳 = まだ.get(q.式);
  if (よい) {
    合++;
    if (訳) 直った.push(q.番 + '  ' + q.式);
  } else if (訳) {
    許した.push(q.番 + '  ' + q.式 + '  => ' + 訳);
  } else {
    違++;
    外れ.push(q.番 + '  ' + q.式 + '  => 台「' + 字 + '」／紙「' + (q.字 || q.答) + '」');
  }
}
console.log('  ★★合った ' + 合 + '★★ ／ 違った ' + 違 + ' ／ ★名指しで まだ ' + 許した.length + '★'
  + ' ／ ★台が 知らない 関数 ' + 押さない + '行★ ／ ★実Excel が 打てなかった ' + 打てない + '行★');
for (const s of 外れ) console.log('       ・' + s);
for (const s of 許した) console.log('       （まだ）' + s);
T('★★押した 行が 全部 紙と 合う★★', 違 === 0, '違った ' + 違 + '行');
T('★★合ったのに「まだ」の ままの 物が 無い★★', 直った.length === 0,
  '直った ' + 直った.length + '本 ... ' + 直った.join(' ') + '\n       ★許しを 外して ください★');
T('★押した 行が 1行でも 在る★', 合 + 違 > 0);
/* ══ ★★合った 数を 決め打ちに する★★ ══（2026-09-18・経営者1 の 型）
     ★増えたら 赤★ ... 「まだ」を 外し忘れて いないか
     ★★減っても 赤★★ ... ★良く なった 時だけ 動く 門は 甘い★
       ＝★今日 私が 経営者1 に 言った「緩い 上限は 門では ない」の 裏返し★
     ⇒★数が 動いたら ★その場で 見る★＝黙って 通さない★ */
const 合った本数 = 60;   /* ★2026-09-19 ... 58 -> 59（ODDL）-> ★60★（★ODDF を 出した＝保留 0個★）★ */
T('★★合った 数が ' + 合った本数 + '★★（★増えても 減っても 赤★）',
  合 === 合った本数,
  '出た ' + 合 + ' ／ 決め打ち ' + 合った本数
  + '（★増えた＝「まだ」を 外して ください ／ 減った＝何かが 壊れました★）');

console.log('');
console.log('  ★★この 試験が 言えない 事★★');
console.log('    ・★台が 知らない 関数の 行（' + 押さない + '行）は 押して いません★');
console.log('      ＝LAMBDA／MAP／REDUCE／SCAN／BYROW／BYCOL／MAKEARRAY など');
console.log('      ＝★書いたら ここが 自動で 増えます★');
console.log('    ・★実Excel が 打てなかった 行（' + 打てない + '行）も 押して いません★');
console.log('      ＝`=LET(x,1)` ... ★誤りの 値では なく「打てない」★＝★5つ目の 型★');
console.log('    ・★画面では ありません★（台だけ）');
console.log('');
console.log('kansuu46-9wakume: ' + pass + ' 緑 / ' + fail + ' 赤');
process.exit(fail ? 1 : 0);
