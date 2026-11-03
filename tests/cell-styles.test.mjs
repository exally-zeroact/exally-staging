/* cell-styles.test.mjs — ★セルのスタイル（ホーム→スタイル）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 を COM で 読んだ）★
 *    Excelの 標準スタイルは ★47個★。よく使う 9つを 中身ごと 測った。
 *    ★COM の Color は BGR★（0xBBGGRR）なので RGB に 直す。
 *      見出し 1  太字 15pt 字 #0E2841 ／ 良い 字 #006100 塗り #C6EFCE
 *      悪い 字 #9C0006 塗り #FFC7CE ／ どちらでもない 字 #9C5700 塗り #FFEB9C
 *      メモ 塗り #FFFFCC ／ 集計 太字 11pt
 *
 *  ★色は 実Excelの 値を そのまま★＝同じ書類を 開いた時に 同じに 見える為（相互運用）。
 *  うちの緑は ★画面の色★。セルの中身の 色は ★ファイルの中身★＝別の話。
 *
 *  走らせ方: node tests/cell-styles.test.mjs [--self-test]
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

const C = require_(path.join(ROOT, 'lib/cell-styles.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

console.log('\n[① ★BGR → RGB が 実Excelの 値と 合う★]');
for (const [bgrHex, 期待, 名] of [
  ['41280E', '#0E2841', '見出しの字'], ['006100', '#006100', '良いの字'],
  ['CEEFC6', '#C6EFCE', '良いの塗り'], ['06009C', '#9C0006', '悪いの字'],
  ['CEC7FF', '#FFC7CE', '悪いの塗り'], ['00579C', '#9C5700', 'どちらでもないの字'],
  ['9CEBFF', '#FFEB9C', 'どちらでもないの塗り'], ['CCFFFF', '#FFFFCC', 'メモの塗り'],
]) {
  ok(名 + ' … BGR ' + bgrHex + ' → ' + 期待, C._bgr(parseInt(bgrHex, 16)) === 期待, C._bgr(parseInt(bgrHex, 16)));
}

console.log('\n[② 一覧の 中身が 実測どおり]');
/* ★2026-08-31：★よ9個 → 47個★（司さん「細胞レベルまで 網羅して 持ち込む」）
     実 Excel の 組み込みは 47個。新規の 空ブックで
     ★スタイルを 1つずつ ★別のセル★に 当てて 読んだ★。
     ★前の 9個と 食い違い 0件★を 確かめてから 広げている。 */
ok('★47個 在る（実 Excel の 組み込み 全部）★', C.一覧.length === 47, String(C.一覧.length));
ok('★先頭は 標準（当てると 元に 戻る）★', C.一覧[0].名 === '標準', C.一覧[0].名);
ok('★英語のままの 名前を 画面の 字に そろえている★',
  !!C.探す('通貨') && !!C.探す('パーセント') && !!C.探す('桁区切り'),
  '通貨/パーセント/桁区切り のどれかが 無い');
ok('★元の 英語名も 残している（ファイルに 書く時に 要る）★',
  C.探す('通貨').英 === 'Currency' && C.探す('標準').英 === 'Normal',
  JSON.stringify([C.探す('通貨').英, C.探す('標準').英]));
ok('★アクセントは 24個（20%/40%/60%/素の ×6）★',
  C.一覧.filter((v) => /アクセント/.test(v.名)).length === 24,
  String(C.一覧.filter((v) => /アクセント/.test(v.名)).length));
ok('★塗りが 在るのは 32個（実測）★',
  C.一覧.filter((v) => v.形.bg).length === 32,
  String(C.一覧.filter((v) => v.形.bg).length));
ok('★色は 全部 #RRGGBB（BGR のまま 残っていない）★',
  C.一覧.every((v) => (!v.形.bg || /^#[0-9A-F]{6}$/.test(v.形.bg))
    && (!v.形.color || /^#[0-9A-F]{6}$/.test(v.形.color))), '色の 形が 違う');
ok('★手で 書いていない（自動生成の 印が 在る）★',
  /自動生成/.test(fs.readFileSync(new URL('../lib/cell-styles.js', import.meta.url), 'utf8')),
  '★手で 書くと 実 Excel と ずれる★');
for (const [名, 期待] of [
  ['見出し 1', { bold: true, fontSize: 15, color: '#0E2841' }],
  ['見出し 2', { bold: true, fontSize: 13, color: '#0E2841' }],
  ['タイトル', { fontSize: 18, color: '#0E2841' }],
  ['良い', { fontSize: 11, color: '#006100', bg: '#C6EFCE' }],
  ['悪い', { fontSize: 11, color: '#9C0006', bg: '#FFC7CE' }],
  ['集計', { bold: true, fontSize: 11, color: '#000000' }],
]) {
  const st = C.探す(名);
  ok(名, st && JSON.stringify(st.形) === JSON.stringify(期待), st ? JSON.stringify(st.形) : '無い');
}
ok('★「標準」は 何も 持たない（元に戻す用）★', JSON.stringify(C.探す('標準').形) === '{}');

console.log('\n[③ 当てる＝触る所だけ 変える]');
{
  const cell = { v: '中身', d: '中身', align: 'right', numFmt: '#,##0' };
  const 控え = C.当てる(cell, '良い');
  ok('字の色が 付く', cell.color === '#006100', String(cell.color));
  ok('塗りが 付く', cell.bg === '#C6EFCE', String(cell.bg));
  ok('★中身は 変えない★', cell.v === '中身', String(cell.v));
  ok('★揃えは 残す（触らない所）★', cell.align === 'right', String(cell.align));
  ok('★書式も 残す★', cell.numFmt === '#,##0', String(cell.numFmt));
  ok('★元に戻す控えを 返す★', Array.isArray(控え) && 控え.length >= 2, JSON.stringify(控え));
}
{
  const cell = { bold: true, fontSize: 15, color: '#0E2841', bg: '#C6EFCE', align: 'left' };
  C.当てる(cell, '標準');
  ok('★「標準」で 4つとも 外れる★',
    cell.bold === undefined && cell.fontSize === undefined && cell.color === undefined && cell.bg === undefined,
    JSON.stringify(cell));
  ok('★揃えは 残る★', cell.align === 'left');
}
{
  ok('★知らない名前は 何も しない★', C.当てる({}, 'そんなスタイルは無い') === null);
}

console.log('\n[④ 画面に 繋がっている]');
ok('部品を 読み込んでいる', /lib\/cell-styles\.js/.test(book));
ok('窓が 在る', /id="styleOverlay"/.test(book));
ok('当てる働きが 在る', /function セルのスタイルを当てる\(/.test(book));
ok('★元に戻せる（控えを 積む）★', /undoStack\.push\(\{ type: 'format', snap: snap \}\)/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
{
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { セルのスタイルを開く: function () { 受け = 'ok'; } };
  ACT['セルのスタイル']();
  g.window = 前w;
  ok('「セルのスタイル」→ セルのスタイルを開く', 受け === 'ok', String(受け));
}

console.log('\ncell-styles: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① BGR を そのまま RGB として 使う（色が 逆に なる） */
  const 素朴 = (n) => '#' + n.toString(16).toUpperCase().padStart(6, '0');
  if (素朴(parseInt('41280E', 16)) === '#0E2841') { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  /* 壊し② 中身まで 変える */
  const cell = { v: 'x' };
  C.当てる(cell, '良い');
  if (cell.v !== 'x') { 素通り++; console.log('  ★素通り★ 中身を 変えた'); }
  /* 壊し③ 「標準」が 何か 持っている */
  if (Object.keys(C.探す('標準').形).length) { 素通り++; console.log('  ★素通り★ 標準が 何か 持っている'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
