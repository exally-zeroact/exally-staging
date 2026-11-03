/* formula-tab.test.mjs — ★数式タブ（計算方法・名前・関数の一覧）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-formula-tab.ps1
 *    ・計算方法の 既定 = ★-4105（自動）★（手動 -4135／データテーブル以外 自動 2）
 *    ・A1:B3 を 選んで「上端行」から 名前を 作ると
 *        月 → =Sheet1!$A$2:$A$3 ／ 売上 → =Sheet1!$B$2:$B$3
 *        ＝★見出しの セルは 中に 入らない★
 *    ・名前に 出来ない字（実測）… 「売 上」（空白）／「1月」（数で始まる）／「A1」（番地の形）
 *      「あ_い」「あ.い」は ★入る★
 *
 *  走らせ方: node tests/formula-tab.test.mjs [--self-test]
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
const NR = require_(path.join(ROOT, 'lib/named-ranges.js'));

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

console.log('\n[① 測った 紙が 残っている]');
ok('tools/measure-formula-tab.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-formula-tab.ps1')));
ok('実測の 値が 見出しに 在る（-4105）', /-4105/.test(book), '');

console.log('\n[② 計算方法（実測＝既定は 自動）]');
for (const n of ['計算方法の窓を開く', '計算方法を決める', '計算していない印']) ok(n + ' が 在る', !!抜く(n));
ok('★はじめは 自動★', /var 計算方法 = 'auto';/.test(book));
ok('窓が 在る', /id="calcOverlay"/.test(book));
ok('★手動なら 波及させない★', /計算方法 === 'manual'[\s\S]{0,120}計算していない印\(true\)/.test(book));
ok('★手動でも 自動でも 通る道が 在る（else）★',
  /計算していない印\(true\);\s*\} else \{[\s\S]{0,120}_scheduleRecalc/.test(book));
ok('★F9 で すべて 計算★', /if\(ek === 'f9'\)/.test(book) && /すべて再計算\(\);/.test(book));
ok('★Shift\+F9 は このシートだけ★', /e\.shiftKey\) このシートを再計算\(\)/.test(book));
ok('★計算したら 印を 消す★', /すべて再計算[\s\S]{0,200}計算していない印\(false\)/.test(book));
ok('★「計算していません」を 画面に 出す★', /計算していません/.test(book));
ok('★出していない物は 出していないと 書く（データテーブル以外 自動）★',
  /実Excelの「データ テーブル以外 自動」は 出していません/.test(book));

console.log('\n[③ 選択範囲から 名前を 作る（実測どおり 見出しは 入らない）]');
{
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2', 'cellAddr',
    抜く('範囲から名前の下ごしらえ') + '\nreturn 範囲から名前の下ごしらえ;');
  const 名 = (r, c) => String.fromCharCode(65 + c) + (r + 1);
  const データ = {
    '0,0': { v: '月' }, '0,1': { v: '売上' },
    '1,0': { v: '1月' }, '1,1': { v: 100 },
    '2,0': { v: '2月' }, '2,1': { v: 150 },
  };
  const 下 = f([{ data: データ, name: 'Sheet1' }], 0, 0, 2, 0, 1, 名);
  const 上から = 下(true, false);
  ok('上端行 … 2つ 作る', 上から.length === 2, JSON.stringify(上から));
  ok('★月 → $A$2:$A$3（見出しの A1 は 入らない）★',
    上から[0].名 === '月' && /A2:A3$/.test(上から[0].参照先), JSON.stringify(上から[0]));
  ok('★売上 → $B$2:$B$3★',
    上から[1].名 === '売上' && /B2:B3$/.test(上から[1].参照先), JSON.stringify(上から[1]));
  const 左から = 下(false, true);
  ok('左端列 … 3つ 作る（見出しの行も 1本に なる）', 左から.length === 3, JSON.stringify(左から.map((x) => x.名)));
  ok('★1月 → B2:B2（左の 見出しは 入らない）★',
    左から[1].名 === '1月' && /B2:B2$/.test(左から[1].参照先), JSON.stringify(左から[1]));
  const 両方 = 下(true, true);
  ok('★両方 ✓ なら 角を 二重に 使わない★',
    両方.length === 3 && 両方[0].名 === '売上', JSON.stringify(両方.map((x) => x.名)));
}

console.log('\n[④ 名前に 出来ない字（実測）]');
for (const [名, 通る, なぜ] of [
  ['売上', true, ''],
  ['あ_い', true, '下線'],
  ['あ.い', true, '点'],
  ['売 上', false, '★空白★'],
  ['1月', false, '★数で 始まる★'],
  ['A1', false, '★番地の形★'],
]) {
  const r = NR.名前を確かめる(名);
  ok((通る ? '入る' : '★入らない★') + ' … ' + 名 + (なぜ ? '  ' + なぜ : ''),
    通る ? r === null : r !== null, String(r));
}
ok('★出来ない物を 黙って とばさない（数を 出す）★', /' \+ とばした\.length \+ '個は 名前に 出来ません/.test(book));
ok('★名前に 出来ない字ばかりなら はじめから ✓ を 付けない★',
  /名にできる\(r1, c\)/.test(book) && /名にできる\(r, c1\)/.test(book));
ok('★作る前に 何が 出来るか 見せる★', !!抜く('範囲から名前の見本'));

console.log('\n[⑤ 関数の 一覧（無い関数を 出さない）]');
for (const n of ['関数の窓を開く', '関数を入れる', '関数を覚える', 'うちに在る関数', '数式で使用の窓を開く',
  '名前を式に入れる']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="funcOverlay"/.test(book));
{
  const f = new Function(抜く('関数を覚える') + '\nvar 最近使った関数 = [];\nreturn function(n){ return 関数を覚える(n); };');
  /* 覚える所は 外の 変数を 使うので、ここでは 中身の 決まりだけ 見る */
  ok('★同じ物を 2回 覚えない（前に 出す）★', /var i = 最近使った関数\.indexOf\(s\);[\s\S]{0,80}splice\(i, 1\)/.test(book));
  ok('★10個までしか 覚えない（実Excelと 同じ）★', /最近使った関数\.length > 10/.test(book));
  ok('★新しい物が 先頭★', /最近使った関数\.unshift\(s\)/.test(book));
}
ok('★うちに 本当に 在る 関数だけ 出す★', /うちに在る関数\(財務の関数\)/.test(book));
ok('  無かった時の 言い方も 在る', /財務の関数が 1つも 入っていません/.test(book));
ok('★押すと 「=関数(」まで 入る★', /setCell\(selR1, selC1, '=' \+ 名 \+ '\('\)/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名, 期待) => {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function (a) { 受け = (a === undefined ? 'ok' : a); };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名 + (期待 === 'ok' ? '' : '（' + 期待 + '）'), 受け === 期待, String(受け));
  };
  試す('計算方法の設定', '計算方法の窓を開く', 'ok');
  試す('最近使った関数', '関数の窓を開く', '最近');
  試す('財務', '関数の窓を開く', '財務');
  試す('数式で使用', '数式で使用の窓を開く', 'ok');
  試す('選択範囲から作成', '範囲から名前の窓を開く', 'ok');
}

console.log('\n[⑦ ワークシート分析（トレース・数式の検証・ウォッチ）]');
{
  for (const n of ['式の中の番地', '参照元のトレース', '参照先のトレース', 'トレース矢印の削除',
    'トレースを描く', '数式の検証を開く', '数式の検証を1手すすめる', 'ウォッチを開く',
    'ウォッチに足す', 'ウォッチから外す', 'ウォッチを描く']) ok(n + ' が 在る', !!抜く(n));
  ok('窓が 在る（検証）', /id="evalOverlay"/.test(book));
  ok('窓が 在る（ウォッチ）', /id="watchOverlay"/.test(book));
  ok('★矢印は 描く所から 呼ぶ★', /トレースを描く\(\);/.test(book));

  const f = new Function('addrToRC', 抜く('式の中の番地') + '\nreturn 式の中の番地;')((a) => {
    const m = a.match(/^([A-Z]+)(\d+)$/i);
    if (!m) return null;
    let col = 0;
    for (let i = 0; i < m[1].length; i++) col = col * 26 + (m[1].toUpperCase().charCodeAt(i) - 64);
    return { r: parseInt(m[2]) - 1, c: col - 1 };
  });
  const 見 = (式) => JSON.stringify(f(式).map((x) => x.r + ',' + x.c));
  ok('=A1+B2 → 2つ', 見('=A1+B2') === '["0,0","1,1"]', 見('=A1+B2'));
  ok('★範囲は 中を 全部 拾う（=SUM(A1:A3)）★', 見('=SUM(A1:A3)') === '["0,0","1,0","2,0"]', 見('=SUM(A1:A3)'));
  ok('★$ が 付いても 拾う★', 見('=$A$1') === '["0,0"]', 見('=$A$1'));
  ok('★同じ番地を 2回 数えない★', 見('=A1+A1') === '["0,0"]', 見('=A1+A1'));
  ok('★"文字" の 中の A1 は 拾わない★', 見('="A1"&A2') === '["1,0"]', 見('="A1"&A2'));
  ok('式でなければ 0個', 見('ただの字') === '[]', 見('ただの字'));
  ok('番地が 無ければ 0個', 見('=1+2') === '[]', 見('=1+2'));

  ok('★見ている セルが 無い時は 理由を 出す★', /このセルは 他のセルを 見ていません/.test(book));
  ok('★見ている セルが 居ない時も 理由を 出す★', /このセルを 見ている セルは ありません/.test(book));
  ok('★式でなければ 検証を 開かない★', /このセルは 式では ありません/.test(book));
  ok('★空セルは 0 として 見せる（実Excelと 同じ）★', /空セルは 0＝Excelと 同じ/.test(book));
  ok('★範囲は 1つの かたまりで 置き換える（{10;20}）★',
    /_検証の範囲の値\(左上, 右下\)/.test(book) && /'\{' \+ 行たち\.join\(';'\) \+ '\}'/.test(book));
  ok('★最後は このセルの 本当の 答えを 出す★',
    /答えは このセルの 本当の 答え/.test(book));
  ok('★参照元と 参照先で 色を 変える★', /a\.種類 === '参照元' \? '#1E88E5' : '#C0392B'/.test(book));
  ok('★計算し直したら ウォッチも 描き直す★', /すべて再計算[\s\S]{0,300}ウォッチを描く\(\)/.test(book));

  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['参照元のトレース', '参照元のトレース'], ['参照先のトレース', '参照先のトレース'],
    ['トレース矢印の削除', 'トレース矢印の削除'], ['数式の検証', '数式の検証を開く'],
    ['ウォッチウィンドウ', 'ウォッチを開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nformula-tab: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 見出しを 中に 入れてしまったら 気づけるか */
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2', 'cellAddr',
    抜く('範囲から名前の下ごしらえ') + '\nreturn 範囲から名前の下ごしらえ;');
  const 名 = (r, c) => String.fromCharCode(65 + c) + (r + 1);
  const 下 = f([{ data: { '0,0': { v: '月' }, '1,0': { v: 1 } }, name: 'S' }], 0, 0, 1, 0, 0, 名);
  const 出 = 下(true, false);
  if (/A1:/.test(出[0].参照先)) { 素通り++; console.log('  ★素通り★ 見出しの A1 が 中に 入っている'); }
  else console.log('  ok   見出しは 入っていない（' + 出[0].参照先 + '）');
  /* 壊し② 名前の 決まりが ゆるくなっていないか */
  if (NR.名前を確かめる('1月') === null) { 素通り++; console.log('  ★素通り★ 数で始まる名前が 通った'); }
  else console.log('  ok   数で始まる名前は 通らない');
  if (NR.名前を確かめる('売 上') === null) { 素通り++; console.log('  ★素通り★ 空白入りが 通った'); }
  else console.log('  ok   空白入りは 通らない');
  /* 壊し③ 手動の 道が 消えていないか */
  if (!/計算方法 === 'manual'/.test(book)) { 素通り++; console.log('  ★素通り★ 手動の 分かれ道が 無い'); }
  else console.log('  ok   手動の 分かれ道が 在る');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
