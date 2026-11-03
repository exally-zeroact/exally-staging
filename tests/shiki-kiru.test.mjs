/* shiki-kiru.test.mjs — ★式を「字の かたまり」に 切る★（2026-09-11）
 *
 *  ★★なぜ 作ったか★★
 *    今 ★式を 読む 所は 借り物（HyperFormula）★です。
 *    司さん「借りんで ええように」＝★土台も 自分で 作る★
 *    土台は 5つ … ①読む ②頼りの 地図 ③計算し直す 順番 ④計算 ⑤溢れ
 *    ⇒★これは ①の 一番 下＝字に 切る 所★（最初の 1枚）
 *
 *  ★★借り物の 中は 1文字も 読んで いません★★
 *    形は ★実Excel が 受ける 字★から 決めました。
 *
 *  ★★測り方＝切って 繋ぎ直したら 元と 1バイトも 違わない事★★
 *    ⇒★取りこぼしが 0★の 証し（実Excel を 呼ばずに 測れる）
 *    実測 … 司さんの 実物 ★15,799本 全部★ 切れて 戻った
 *
 *  ★この 台は エンジンを 建てません★＝字を 切るだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/shiki-kiru.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[shiki-kiru] ★式を 字の かたまりに 切る★');

/* ★実Excel が 受ける 形★（実物に 無い 物も 入れる＝「在るから 安心」に しない） */
const 試 = [
  '=1+2*3', '=-5', '=+5', '=1.5', '=.5', '=1.64E-14', '=1E+20', '=2E5',
  '="あ"&"い"', '="引用""符"', '=A1', '=$A$1', '=A1:B2', '=A:A', '=1:1',
  '=Sheet1!A1', "='歩合 表'!A1", "='それ''これ'!A1", '=SUM(A1:A3)', '=SUM(A1:A3,B1)',
  '=IF(A1<0,"負","正")', '=A1<>B1', '=A1>=B1', '=A1<=B1', '=10%', '=2^3',
  '=#N/A', '=IFERROR(A1,#N/A)', '=表[列]', '=表[[#見出し],[列]]', '={1,2;3,4}',
  '=SUM(A1:A3 B1:B3)', '=IF(TRUE,1,0)', '=A1&""',
  '=SUM((A1:A3)*(B1:B3))', '=INDEX(表[見本一四],MATCH(B4,表[日付],0))',
];

T('★切って 繋ぎ直すと 元に 戻る（' + 試.length + '通り）★', () => {
  const 違い = [];
  for (const t of 試) {
    const r = K.切る(t);
    if (!r.ok) { 違い.push(t + ' … ' + r.なぜ); continue; }
    const 戻 = K.繋ぐ(r.出);
    if (戻 !== t) 違い.push(t + ' → ' + 戻);
  }
  if (違い.length) throw new Error('★' + 違い.length + '本 違う★  ' + 違い.slice(0, 3).join(' ／ '));
});

T('★★かたまりの 中身が 正しい（数を 焼き込まない）★★', () => {
  /* ★「何個に 切れたか」は 書きません★＝切り方を 良くしたら 変わる 数です。
     ★見るのは「この 字が 1つの かたまりに なって いるか」★ */
  const 見る = [
    ['="引用""符"', '字', '"引用""符"'],
    ["='歩合 表'!A1", '名', "'歩合 表'!A1"],
    ['=#N/A', '誤', '#N/A'],
    ['=1.64E-14', '数', '1.64E-14'],
    ['=表[[#見出し],[列]]', '名', '表[[#見出し],[列]]'],
    ['=A1<>B1', '記', '<>'],
  ];
  const 違い = [];
  for (const [式, 型, 字] of 見る) {
    const r = K.切る(式);
    if (!r.ok) { 違い.push(式 + ' … ' + r.なぜ); continue; }
    const 在る = r.出.some((x) => x.型 === 型 && x.字 === 字);
    if (!在る) 違い.push(式 + ' に 「' + 字 + '」（' + 型 + '）が 無い … ' + JSON.stringify(r.出.map((x) => x.字)));
  }
  if (違い.length) throw new Error('★' + 違い.length + '本★  ' + 違い.join(' ／ '));
});

T('★★壊れた 式は ★断る★（勝手に 直さない）★★', () => {
  const 壊 = ['="閉じない', "='閉じない", '=表[閉じない'];
  const 通った = [];
  for (const t of 壊) { if (K.切る(t).ok) 通った.push(t); }
  if (通った.length) {
    throw new Error('★' + 通った.length + '本 通して しまった★  ' + 通った.join(' ／ ')
      + '  ⇒★壊すより 断る★');
  }
});

T('★★読めない 字が 来たら 断る★★', () => {
  /* ★「読めた事に する」のが 一番 危ない★＝黙って 落ちた 字が 出る */
  const r = K.切る('=A1' + String.fromCharCode(0) + 'B1');
  if (r.ok) throw new Error('★読めない 字を 通した★');
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
