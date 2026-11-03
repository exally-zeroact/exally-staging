/* view-window.test.mjs — ★表示タブ（ブックの表示・シートビュー・ウィンドウ・分割）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-view.ps1 / tools/measure-sheetview.ps1
 *    Window.View … ★1=標準／2=改ページ プレビュー／3=ページ レイアウト★
 *    固定 … C3 で SplitRow=2 SplitColumn=2（先頭行だけ 1/0・先頭列だけ 0/1）
 *    分割 … C3 で SplitRow=2 SplitColumn=2・★Panes=4★／外すと 1
 *    ズーム … 選択範囲に 合わせる＝A1:E10 で ★243★・A1:B3 で ★400（上限）★
 *    枠線/見出し … 見せる=True・見せる=True／★刷る=False・刷る=False★
 *    改ページ … `HPageBreaks.Add(A20)` → ★20行目で 切れる★
 *    ユーザー設定のビュー … 隠した行と 倍率を 覚え `Show()` で ★3行目 隠れ=True・Zoom=75★ に 戻る
 *    シート ビュー … Count 0 → Add('名') で 1 ／ Exit() で 抜ける
 *    窓の 名前 … 1つの時「Book5」／2つに すると ★「Book5  -  2」「Book5  -  1」★（空白2つ）
 *    表示しない … Visible=False。★窓の 数は 減らない★
 *    並べて比較 … ★上下★に 並ぶ（1つ目 上0.2／2つ目 上286.8・幅は 同じ）
 *
 *  走らせ方: node tests/view-window.test.mjs [--self-test]
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
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const VM = require_(path.join(ROOT, 'lib/view-mode.js'));
const SV = require_(path.join(ROOT, 'lib/sheet-view.js'));

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
ok('tools/measure-view.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-view.ps1')));
ok('tools/measure-sheetview.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-sheetview.ps1')));

console.log('\n[② 紙の 大きさ（実測の 余白・用紙から）]');
{
  const n = VM.紙の中身({});
  ok('★A4縦の 中身は 494.5 × 733.9 点★',
    Math.abs(n.幅 - 494.48) < 0.1 && Math.abs(n.高 - 733.89) < 0.1,
    n.幅.toFixed(2) + ' x ' + n.高.toFixed(2));
  const y = VM.紙の中身({ 向き: 'landscape' });
  /* ★実測（08-30）★＝実Excel は 向きを 変えても 余白を 入れ替えない
     （縦 左右50.4/上下54 → 横も 左右50.4/上下54）。
     だから 中身は ★紙だけ 入れ替えて 余白は そのまま★＝単純な 入れ替えには ならない。 */
  ok('★横向きは 紙だけ 入れ替え・余白は そのまま（実測）★',
    Math.abs(y.幅 - (29.7 - 1.778 * 2) * VM.一cmの点) < 0.01 &&
    Math.abs(y.高 - (21.0 - 1.905 * 2) * VM.一cmの点) < 0.01,
    y.幅.toFixed(2) + ' x ' + y.高.toFixed(2));
  ok('  横の方が 幅が 広い', y.幅 > n.幅 && y.高 < n.高);
  ok('★倍率50％は 入る量が 2倍★', Math.abs(VM.紙の中身({ 倍率: 50 }).高 - n.高 * 2) < 0.01);
  ok('★倍率は 10〜400 で 止める★',
    Math.abs(VM.紙の中身({ 倍率: 1000 }).高 - VM.紙の中身({ 倍率: 400 }).高) < 0.01 &&
    Math.abs(VM.紙の中身({ 倍率: 1 }).高 - VM.紙の中身({ 倍率: 10 }).高) < 0.01);
  ok('★余白の 既定は 実測（上下1.905 / 左右1.778）★',
    VM.既定の余白.上 === 1.905 && VM.既定の余白.左 === 1.778);
  ok('★1cm は 28.3465点★', Math.abs(VM.一cmの点 - 28.3464567) < 1e-6);
}

console.log('\n[③ ★切れ目★（列は 実Excel と 同じ 8列に なるか）]');
{
  const n = VM.紙の中身({});
  const 列幅 = Array(20).fill(80 * 0.75);      /* うちの 既定 80点 ＝ 紙の 60点 */
  const cb = VM.切れ目(列幅, n.幅, []);
  ok('★1ページ 8列（実Excel の 実測と 同じ）★', VM.一ページの数(cb, 20) === 8, JSON.stringify(cb));
  ok('★切れ目は 9列目・17列目（0から 数えて 8・16）★',
    cb[0] === 8 && cb[1] === 16, JSON.stringify(cb));
  const 行高 = Array(120).fill(22 * 0.75);
  const rb = VM.切れ目(行高, n.高, []);
  /* ★違い★＝実Excel は 39行（プリンタの 刷れない縁）。うちは 紙から 数えて 44行。 */
  ok('★うちの 行は 44行（22点の 行で 紙から 数えた）★', VM.一ページの数(rb, 120) === 44,
    String(VM.一ページの数(rb, 120)));
  ok('★手で 足した 所で 切れる（実測＝A20 は 20行目で 切れる）★',
    VM.切れ目(行高, n.高, [19])[0] === 19, JSON.stringify(VM.切れ目(行高, n.高, [19]).slice(0, 3)));
  ok('★同じ 番号を 2回 出さない★',
    VM.切れ目(行高, n.高, [44]).filter(v => v === 44).length === 1,
    JSON.stringify(VM.切れ目(行高, n.高, [44]).slice(0, 3)));
  ok('★ページ数＝(行の切れ目+1)×(列の切れ目+1)★', VM.ページ数(rb, cb) === (rb.length + 1) * 3,
    String(VM.ページ数(rb, cb)));
  ok('★切れ目が 無ければ 1ページ★', VM.ページ数([], []) === 1);
}

console.log('\n[④ ブックの 表示（1/2/3＝実Excel と 同じ 番号）]');
ok('★表示の 番号は 1=標準・2=改ページ・3=ページレイアウト★',
  VM.表示たち[0].番 === 1 && VM.表示たち[1].名 === '改ページ プレビュー' && VM.表示たち[2].番 === 3);
for (const n of ['ブックの表示を変える', '標準の表示にする', '改ページプレビューにする',
  'ページレイアウト表示にする', '紙の線を描く', '紙の切れ目', '紙の帯を出す',
  '改ページを足す', '改ページを外す', '改ページを全部外す', '改ページの窓を開く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★1・2・3 以外は 断る★', /1・2・3 の どれかです/.test(book));
ok('★標準（1）の 時は 紙の線を 描かない★', /if \(ブックの表示 === 1\) return 0;/.test(book));
ok('★描くのは 固定した帯の 後★', /紙の線を描く\(\);\n\s*\/\* ★分割/.test(book) ||
  book.indexOf('紙の線を描く();') > book.indexOf('★固定した所の境目に線を1本★'));
ok('★A1 には 切れ目を 置けない★', /A1 の 所には 切れ目を 置けません/.test(book));
ok('★「機械で 変わる」と 画面に 書いてある★', /1ページに 入る 数は 印刷する 機械で 変わります/.test(book));

console.log('\n[⑤ 分割（実測＝C3 で Panes 4）]');
for (const n of ['分割する', '分割をやめる', '分割の側を描く', '分割の線を描く', '分割の高さ', '分割の幅']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★A1 では 分けられない★', /A1 では 分けられません/.test(book));
ok('★押すたび 入れ替わる（もう一度 押すと やめる）★',
  /if \(分割\.行 \|\| 分割\.列\) return 分割をやめる\(\);/.test(book));
ok('★上の側と 左の側を 別の ずれで 描く（固定は 0・分割は その側の ずれ）★',
  /scrollTop = 分割\.上のずれ;/.test(book) && /scrollLeft = 分割\.左のずれ;/.test(book));
ok('★指が 居る 側だけ 動かす★', /var 上の側 = 分割\.行 > 0 && _p\.y < HDR_H \+ 分割の高さ\(\);/.test(book));
ok('★固定（freeze）は 0 の まま（分割と 混ぜない）★',
  /scrollTop=0; wrapH=HDR_H\+_fH; _renderPass\(\);/.test(book));

console.log('\n[⑥ シート ビュー／ユーザー設定のビュー]');
{
  ok('★同じ 名前を 2つ 作らない★', SV.名前を決める(['ビュー'], 'ビュー') === 'ビュー 2');
  ok('  3つ目は 3', SV.名前を決める(['ビュー', 'ビュー 2'], 'ビュー') === 'ビュー 3');
  const たち = [];
  const いま = { 絞り: { 3: true }, 手で隠した行: { 5: true }, 手で隠した列: {}, 倍率: 0.75, 固定行: 2, 固定列: 1 };
  const v = SV.足す(たち, 'テスト', いま);
  ok('★Add は 足した 物を 返す（実測と 同じ）★', v && v.名 === 'テスト' && たち.length === 1);
  /* ★実測＝隠した行と 倍率を 覚え、変えても 戻る★ */
  いま.手で隠した行[9] = true;
  いま.倍率 = 1;
  const 出 = SV.取り出す(たち, 'テスト');
  ok('★後から 変えても 覚えた物は 変わらない★',
    !出.手で隠した行[9] && 出.倍率 === 0.75, JSON.stringify(出));
  ok('★戻した物を いじっても 元は 壊れない★', (() => {
    出.手で隠した行[7] = true;
    return !SV.取り出す(たち, 'テスト').手で隠した行[7];
  })());
  ok('★消せる★', SV.消す(たち, 'テスト') === true && たち.length === 0);
  ok('★無い名前は false★', SV.消す(たち, 'ない') === false);
  ok('★無い名前は null★', SV.取り出す(たち, 'ない') === null);
}
for (const n of ['シートビューを新規', 'シートビューを保持', 'シートビューを終了',
  'シートビューを切り替え', 'シートビューを消す', 'シートビューの窓を開く',
  'シートビューのオプション', 'ブックのビューの窓を開く', 'ブックのビューを足す',
  'ブックのビューを出す', 'ブックのビューを消す', '_今の見え方', '_見え方を当てる']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★ビューは シートごとに 分けて 持つ★', /function _このシートのビュー\(\)/.test(book));
ok('★「新規」より 先に「保持」は 断る★', /先に「新規」で ビューを 作ってください/.test(book));
ok('★共同編集が 無い事を 画面に 書いてある★', /1人で 使うので/.test(book));

console.log('\n[⑦ ウィンドウ]');
for (const n of ['新しいウィンドウを開く', '窓を整列', '_整列の場所', '整列の窓を開く',
  '並べて比較', '窓の位置を元に戻す', '同時にスクロールを切り替え', '窓を表示しない',
  '窓を再表示', '窓の切り替えを開く', '窓へ移る', '_窓の名前', '_生きている窓']) {
  ok(n + ' が 在る', !!抜く(n));
}
{
  const f = new Function('BookOpen', 抜く('_窓の名前') + '\nreturn _窓の名前;')(undefined);
  ok('★1つの時は 名前だけ（実測 "Book5"）★', f(0) === 'ブック', f(0));
  ok('★2つ 以上は 「名前  -  番号」＝空白 2つ（実測）★', f(2) === 'ブック  -  2', JSON.stringify(f(2)));
  ok('  空白が ちょうど 2つ', /ブック {2}- {2}2/.test(f(2)), JSON.stringify(f(2)));
}
{
  const f = new Function(抜く('_整列の場所') + '\nreturn _整列の場所;')();
  const 上下 = f('上下', 1, 2, 1000, 800);
  ok('★上下＝幅は 画面いっぱい・高さは 半分（実測の 並べて比較と 同じ 形）★',
    上下.幅 === 1000 && 上下.高 === 400 && 上下.上 === 400 && 上下.左 === 0, JSON.stringify(上下));
  const 左右 = f('左右', 1, 2, 1000, 800);
  ok('★左右＝幅が 半分・高さは いっぱい★',
    左右.幅 === 500 && 左右.高 === 800 && 左右.左 === 500, JSON.stringify(左右));
  const 重 = f('重ねて', 2, 3, 1000, 800);
  ok('★重ねて＝30点ずつ ずらす★', 重.左 === 60 && 重.上 === 60, JSON.stringify(重));
  const 並 = f('並べて', 0, 4, 1000, 800);
  ok('★並べて＝4つなら 2×2★', 並.幅 === 500 && 並.高 === 400, JSON.stringify(並));
}
ok('★並べて比較は 上下（実測＝上0.2 と 上286.8）★', /var 数 = 窓を整列\('上下'\);/.test(book));
ok('★窓が 無い時は 断る★', /比べる 相手の 窓が ありません/.test(book));
ok('★窓を 開けなかった 時は そう 言う★', /新しい 窓を 開けませんでした/.test(book));
ok('★ブラウザは 他の 窓を 数えられない事を 書いてある★', /他の アプリの 窓を 数えられません/.test(book));
ok('★消せない事を 正直に 書いてある★', /ブラウザは 窓を 消せません/.test(book));
ok('★同時にスクロールは 合図で 合わせる★', /new BroadcastChannel\('exally-book-windows'\)/.test(book));
ok('★スクロールしたら 知らせる★', /if \(同時にスクロール\) _スクロールを知らせる\(\);/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑧ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑧-2 ★下の 帯が シートのタブに 隠れていないか★]');
{
  /* 08-29 に #headerSay で 同じ罠を 踏み、08-30 に また 踏んだ。
     .bottom-stack は z-index:100 で 下に 貼り付いている ＝
     z-index 60 の 帯は ★「在る」のに 見えない★（elementFromPoint で 見つけた）。 */
  ok('★上に 置く 部品が 1つ 在る★', !!抜く('帯を下の上に置く'));
  ok('★高さを その場で 測る（決め打ちしない）★',
    /getBoundingClientRect\(\)\.height/.test(抜く('帯を下の上に置く') || ''));
  ok('★z は 101 から（.bottom-stack の 100 より 上）★',
    /zIndex = String\(101 \+ i\);/.test(抜く('帯たちを並べる') || ''));
  for (const 帯 of ['紙の帯を出す', 'シートビューの帯を出す', '窓の帯を出す']) {
    const 中 = 抜く(帯) || '';
    ok('  ' + 帯 + ' が 使っている', /帯を下の上に置く\(帯\);/.test(中));
    /* ★出してから 並べる★＝出す前に 並べると その 帯は 数に 入らず
       ★bottom:0 の まま 他の 帯の 下に 潜る★（08-30 実測） */
    ok('  ' + 帯 + ' は ★出してから★ 並べている',
      中.indexOf("display = 'flex'") < 中.indexOf('帯を下の上に置く(帯);'),
      String(中.indexOf("display = 'flex'")) + ' < ' + String(中.indexOf('帯を下の上に置く(帯);')));
  }
  ok('★積む 部品が 在る★', !!抜く('帯たちを並べる'));
  /* 08-30：データタブの connBar を 足したので 4本 */
  ok('★下から paperBar → viewBar → winBar → connBar の 順★',
    /\['paperBar', 'viewBar', 'winBar', 'connBar'\]/.test(抜く('帯たちを並べる') || ''));
  ok('★出ていない 帯は 積まない★',
    /getComputedStyle\(e\)\.display === 'none'\) continue;/.test(抜く('帯たちを並べる') || ''));
  ok('★AIの 札に かぶらないよう 右を 空ける★',
    /paddingRight = \(AIの札の幅\(\) \+ 16\)/.test(抜く('帯たちを並べる') || ''));
  ok('★札は ★左はし★で 測る（幅では ない）★',
    /innerWidth - r\.left/.test(抜く('AIの札の幅') || ''));
  ok('★隠していた 本物は #ai-float-btn★',
    /getElementById\('ai-float-btn'\)/.test(抜く('AIの札の幅') || ''));
  ok('★下に 居ない 時は 押さない★',
    /r\.bottom < innerHeight - 160\) return 0;/.test(抜く('AIの札の幅') || ''));
  /* ★帯を 1本 足したら ここも 増える★＝消す 所で 並べ直さないと 下に 穴が あく
     ★★2026-09-11 … 数を 書き込むのを やめました★★
       ★前★ … `並.length >= 3`＝「3本 以上 在るはず」と ★数を 焼き込んで★ いた
         ⇒ 帯を 2本に 減らしたら ★中身は 正しいのに 赤★ に なる
       ★今★ … ★くらべる 所は そのまま★（帯の 本数 ＝ 消す 所の 数）／
         ★1本も 無い時だけ 赤★＝それは「探す 所が 違う（空振り）」の 合図 */
  {
    const 並 = (抜く('帯たちを並べる') || '').match(/'[a-zA-Z]+Bar'/g) || [];
    const 消 = (book.match(/帯\.style\.display = 'none'; 帯たちを並べる\(\); return 0;/g) || []).length;
    ok('★帯を 1本でも 見つけて いる（空振りして いない）★', 並.length > 0, '帯 ' + 並.length + '本');
    ok('★帯の 本数と「消した時に 並べ直す」所の 数が 合っている★',
      並.length === 消, '帯 ' + 並.length + '本 / 消す所 ' + 消 + 'か所');
  }
}

console.log('\n[⑨ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('シートビュー切替', 'シートビューの窓を開く');
  試す('シートビュー保持', 'シートビューを保持');
  試す('シートビュー終了', 'シートビューを終了');
  試す('シートビュー新規', 'シートビューを新規');
  試す('シートビュー設定', 'シートビューのオプション');
  試す('標準の表示', '標準の表示にする');
  試す('改ページプレビュー', '改ページプレビューにする');
  試す('ページレイアウト表示', 'ページレイアウト表示にする');
  試す('ユーザー設定のビュー', 'ブックのビューの窓を開く');
  試す('新しいウィンドウ', '新しいウィンドウを開く');
  試す('ウィンドウ整列', '整列の窓を開く');
  試す('ウィンドウ分割', '分割する');
  試す('ウィンドウ非表示', '窓を表示しない');
  試す('ウィンドウ再表示', '窓を再表示');
  試す('並べて比較', '並べて比較');
  試す('同時にスクロール', '同時にスクロールを切り替え');
  試す('ウィンドウ位置戻す', '窓の位置を元に戻す');
  試す('ウィンドウ切替', '窓の切り替えを開く');
  試す('改ページ', '改ページの窓を開く');
}

console.log('\nview-window: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const n = VM.紙の中身({});
  const cb = VM.切れ目(Array(20).fill(60), n.幅, []);
  if (VM.一ページの数(cb, 20) !== 8) { 素通り++; console.log('  ★素通り★ 1ページが 8列で ない'); }
  else console.log('  ok   1ページ 8列');
  const たち = [];
  SV.足す(たち, 'あ', { 手で隠した行: { 1: true } });
  const 出 = SV.取り出す(たち, 'あ');
  出.手で隠した行[2] = true;
  if (SV.取り出す(たち, 'あ').手で隠した行[2]) { 素通り++; console.log('  ★素通り★ 取り出した物が 元と つながっている'); }
  else console.log('  ok   取り出した物は 写し');
  const f = new Function(抜く('_整列の場所') + '\nreturn _整列の場所;')();
  if (f('上下', 1, 2, 1000, 800).高 !== 400) { 素通り++; console.log('  ★素通り★ 上下の 高さが 半分で ない'); }
  else console.log('  ok   上下は 半分');
  if (!/if \(ブックの表示 === 1\) return 0;/.test(book)) { 素通り++; console.log('  ★素通り★ 標準でも 紙の線を 描く'); }
  else console.log('  ok   標準では 描かない');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
