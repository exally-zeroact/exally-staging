/* csv-in.test.mjs — ★テキストまたは CSV から★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-csv.ps1
 *    「名,金 / "山田, 太郎",100 / "あ""い",200 / "1行目(改行)2行目",300」を 開くと
 *      1行目: '名' '金'
 *      2行目: ★'山田, 太郎'★ '100'   （"…" の 中の カンマは 区切りに しない）
 *      3行目: ★'あ"い'★ '200'        （"" は 1つの "）
 *      4行目: ★'1行目(改行)2行目'★ '300'（"…" の 中の 改行は 1つの セル）
 *      使った範囲 = ★A1:B4★
 *    Shift_JIS の CSV も 実Excel は そのまま 読める（A2='あいう'）
 *
 *  走らせ方: node tests/csv-in.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};
const C = require_(path.join(ROOT, 'lib/csv-in.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const CR = String.fromCharCode(13), LF = String.fromCharCode(10);

function 抜く(名) {
  const i = book.indexOf('function ' + 名 + '(');
  if (i < 0) return null;
  let d = 0; const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-csv.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-csv.ps1')));

console.log('\n[② ★実測と 同じ 読み方★]');
{
  const 字 = '名,金' + CR + LF + '"山田, 太郎",100' + CR + LF + '"あ""い",200' + CR + LF
    + '"1行目' + CR + LF + '2行目",300' + CR + LF;
  const r = C.読む(字);
  ok('★4行 × 2列（実測 A1:B4）★', r.行 === 4 && r.列 === 2, r.行 + '×' + r.列);
  ok('1行目 … 名／金', r.表[0].join('|') === '名|金', r.表[0].join('|'));
  ok('★"…" の 中の カンマは 区切りに しない★', r.表[1][0] === '山田, 太郎', r.表[1][0]);
  ok('★"" は 1つの "★', r.表[2][0] === 'あ"い', r.表[2][0]);
  ok('★"…" の 中の 改行は 1つの セル★', r.表[3][0] === '1行目' + LF + '2行目',
    JSON.stringify(r.表[3][0]));
  ok('★最後の 改行で 空の 行を 作らない★', r.表.length === 4, String(r.表.length));
}

console.log('\n[③ 区切りを 当てる]');
ok('カンマ', C.区切りを当てる('a,b,c') === ',');
ok('タブ', C.区切りを当てる('a\tb\tc') === '\t');
ok('セミコロン', C.区切りを当てる('a;b;c') === ';');
ok('★"…" の 中の カンマは 数えない★', C.区切りを当てる('"a,b,c,d"\tx\ty') === '\t',
  JSON.stringify(C.区切りを当てる('"a,b,c,d"\tx\ty')));
ok('何も 無ければ カンマ', C.区切りを当てる('abc') === ',');

console.log('\n[④ こまかい所]');
{
  ok('★BOM を 落とす★', C.読む('﻿a,b').表[0][0] === 'a', JSON.stringify(C.読む('﻿a,b').表[0][0]));
  const r = C.読む('a,b,c' + LF + 'x' + LF);
  ok('★行の 長さを そろえる（黙って 落とさない）★',
    r.表[1].length === 3 && r.表[1][1] === '' && r.表[1][2] === '', JSON.stringify(r.表));
  ok('LF だけでも 読む', C.読む('a,b' + LF + 'c,d').行 === 2);
  ok('空の 字なら 0行', C.読む('').行 === 0, String(C.読む('').行));
  ok('★化けを 見つける★', C.化けているか('あ�い') === true);
  ok('★化けていなければ false★', C.化けているか('あいう') === false);
}

console.log('\n[⑤ 画面に つながっている]');
for (const n of ['CSVの窓を開く', 'CSVを読む', 'CSVを入れる', 'データの取得を開く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('入り口が 在る', /id="csvFileInput"/.test(book));
ok('★.csv / .txt / .tsv を 受ける★', /accept="\.csv,\.txt,\.tsv"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/csv-in\.js/.test(book));
ok('★化けたら Shift_JIS で 読み直す★', /TextDecoder\('shift_jis'\)/.test(book));
ok('★それでも 化けていたら そう 言う★', /字が 化けているかもしれません/.test(book));
ok('★新しい シートに 入れる★', /CSVを入れる[\s\S]{0,400}addSheet\(\)/.test(book));
ok('★何行×何列・区切り・字の種類を 言う★', /行 × ' \+ r\.列 \+ '列／/.test(book));
ok('★出来ない 読み口を 並べない★', /まだ 出来ないので 並べていません/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['データの取得', 'データの取得を開く'], ['テキストまたはCSVから', 'CSVの窓を開く'],
    ['最近使ったソース', 'openRireki'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\ncsv-in: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const 字 = '"山田, 太郎",100';
  const r = C.読む(字);
  if (r.表[0].length !== 2) { 素通り++; console.log('  ★素通り★ "…" の 中で 割れている … ' + JSON.stringify(r.表[0])); }
  else console.log('  ok   "…" の 中では 割れない');
  if (C.読む('"あ""い",1').表[0][0] !== 'あ"い') { 素通り++; console.log('  ★素通り★ "" が 1つの " に なっていない'); }
  else console.log('  ok   "" は 1つの "');
  if (!/TextDecoder\('shift_jis'\)/.test(book)) { 素通り++; console.log('  ★素通り★ 読み直す道が 無い'); }
  else console.log('  ok   読み直す道が 在る');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
