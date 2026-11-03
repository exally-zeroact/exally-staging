/* review2.test.mjs — ★校閲（文章校正・見やすさ・翻訳・変更内容・共有・インク）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-review2.ps1
 *    `CheckSpelling` … `recieve`→★False★／`receive`→True／`teh`→★False★／`the`→True／
 *      `apple`→True／`こんにちは`→★True（日本語は いつも 通る）★
 *    ★直しの 案（GetSpellingSuggestions）は COM から 呼べなかった★
 *    代替テキスト … ★はじめは 空★（`''`）／入れられる／`Title` も 在る
 *    共有 … `MultiUserEditing`=★False★／`KeepChangeHistory`=True／
 *      `ChangeHistoryDuration`=★0★／`HighlightChangesOnScreen`=False
 *      ★共有していないと `HighlightChangesOptions` は 通らない（0x800A03EC）★
 *    インク … ★形（Shapes）の 1つとして 数える★
 *    類義語・翻訳 … ★COM から 呼べない★
 *
 *  走らせ方: node tests/review2.test.mjs [--self-test]
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
const R = require_(path.join(ROOT, 'lib/review.js'));

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
ok('tools/measure-review2.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-review2.ps1')));

console.log('\n[② ★実Excel が 測った 6語と 同じ 答えに なるか★]');
{
  /* ★実測★ recieve=False（＝あやしい）／receive=True（＝通る）… */
  const 実測 = [['recieve', false], ['receive', true], ['teh', false], ['the', true],
    ['apple', true], ['こんにちは', true]];
  let 合 = 0;
  for (const [語, 通る] of 実測) {
    const 出 = R.スペルを見る([{ 値: 語, 場所: 'A1' }]);
    const うちも通る = 出.length === 0;
    if (うちも通る === 通る) 合++;
    ok('  ' + 語 + ' … ' + (通る ? '通る' : 'あやしい') + '（実測）',
      うちも通る === 通る, JSON.stringify(出.map(v => v.語)));
  }
  ok('★6語 全部 実Excel と 同じ★', 合 === 6, 合 + '/6');
  ok('★recieve の 直しは receive★',
    R.スペルを見る([{ 値: 'recieve', 場所: 'A1' }])[0].直し === 'receive');
  ok('★teh の 直しは the★', R.スペルを見る([{ 値: 'teh', 場所: 'A1' }])[0].直し === 'the');
  ok('★日本語が 混ざった セルは まるごと 通す★',
    R.スペルを見る([{ 値: 'これは recieve です', 場所: 'A1' }]).length === 0);
  ok('★大文字でも 見つける★', R.スペルを見る([{ 値: 'Recieve', 場所: 'A1' }]).length === 1);
  ok('★場所を 覚えている★',
    R.スペルを見る([{ 値: 'teh', 場所: 'Sheet1!B3' }])[0].場所 === 'Sheet1!B3');
}

console.log('\n[③ ★同じ ブックの 中の 打ちまちがい★（うちの 足し前）]');
{
  const 出 = R.スペルを見る([
    { 値: 'apple', 場所: 'A1' }, { 値: 'apple', 場所: 'A2' },
    { 値: 'apple', 場所: 'A3' }, { 値: 'aplle', 場所: 'A4' }]);
  ok('★3回 出る 言葉に 1文字違いが 1回 → 見つける★',
    出.length === 1 && 出[0].語 === 'aplle' && 出[0].直し === 'apple',
    JSON.stringify(出));
  ok('★2回しか 出ない 言葉では 出さない（うるさく しない）★',
    R.スペルを見る([{ 値: 'apple', 場所: 'A1' }, { 値: 'apple', 場所: 'A2' },
      { 値: 'aplle', 場所: 'A3' }]).length === 0);
  ok('★短い 言葉（3文字以下）は 見ない★',
    R.スペルを見る([{ 値: 'cat', 場所: 'A1' }, { 値: 'cat', 場所: 'A2' },
      { 値: 'cat', 場所: 'A3' }, { 値: 'cut', 場所: 'A4' }]).length === 0);
  ok('★2文字 違うと 出さない★',
    R.スペルを見る([{ 値: 'apple', 場所: 'A1' }, { 値: 'apple', 場所: 'A2' },
      { 値: 'apple', 場所: 'A3' }, { 値: 'axxle', 場所: 'A4' }]).length === 0);
  ok('★同じ 言葉を 2回 出さない（よくある間違い と 重ならない）★',
    R.スペルを見る([{ 値: 'the', 場所: 'A1' }, { 値: 'the', 場所: 'A2' },
      { 値: 'the', 場所: 'A3' }, { 値: 'teh', 場所: 'A4' }])
      .filter(v => v.語 === 'teh').length === 1);
}
console.log('  ─ 1文字ずつの 違いの 数え方 ─');
ok('  同じ = 0', R.違いの数('apple', 'apple') === 0);
ok('  1文字 入れ替え = 1', R.違いの数('aplle', 'apple') === 1);
ok('  1文字 足す = 1', R.違いの数('aple', 'apple') === 1);
ok('  ★長さが 2以上 違えば すぐ 2★', R.違いの数('a', 'apple') === 2);
ok('  大文字小文字は 同じ', R.違いの数('Apple', 'apple') === 0);
ok('  言葉に 分けられる', R.言葉に分ける("don't stop, now!").join(',') === "don't,stop,now");
ok('  数字だけは 拾わない', R.言葉に分ける('123 456').length === 0);

console.log('\n[④ 見やすさ（実測＝代替テキストは はじめ 空）]');
{
  const 出 = R.見やすさを見る({
    物たち: [{ 名: '図1' }, { 名: '図2', 代替: 'まるい図' }],
    結合たち: ['A1:B2'],
    シート名たち: ['Sheet1', '売上', 'シート3'],
    色の組たち: [{ 字: '#CCCCCC', 背: '#FFFFFF', 場所: 'C3' },
      { 字: '#000000', 背: '#FFFFFF', 場所: 'C4' }],
  });
  ok('★説明の 無い 図だけ 赤★', 出.filter(v => /説明が 無い/.test(v.何)).length === 1);
  ok('  結合を 見つける', 出.filter(v => /結合/.test(v.何)).length === 1);
  ok('★Sheet1 と シート3 の 2つを 見つける★',
    出.filter(v => /名前が そのまま/.test(v.何)).length === 2);
  ok('  「売上」は 出さない', !出.some(v => v.どこ === '売上'));
  ok('★色が 近い 所だけ 出す★', 出.filter(v => /色が 近すぎる/.test(v.何)).length === 1);
  ok('  黒×白は 出さない', !出.some(v => v.どこ === 'C4'));
  ok('  何も 無ければ 0個', R.見やすさを見る({}).length === 0);
}
console.log('  ─ 明るさの 比（見やすさの 決まりの 数え方）─');
ok('★黒と白は 21.0:1★', Math.abs(R.明るさの比('#000000', '#FFFFFF') - 21) < 0.01,
  R.明るさの比('#000000', '#FFFFFF').toFixed(2));
ok('★同じ 色は 1.0:1★', Math.abs(R.明るさの比('#777777', '#777777') - 1) < 0.001);
ok('  順は 関係ない',
  Math.abs(R.明るさの比('#000000', '#FFFFFF') - R.明るさの比('#FFFFFF', '#000000')) < 1e-9);
ok('  3文字の 色も 読める', Math.abs(R.明るさの比('#000', '#fff') - 21) < 0.01);

console.log('\n[⑤ 画面から 押せる]');
for (const n of ['スペルを見る', 'スペルを直す', '見やすさを見る', '絵の説明を入れる',
  '絵の説明を決める', '類義語を聞く', '翻訳を聞く', '_AIに聞く', 'AIへ送る',
  '変更内容を見る', '共有を解除', 'インクを隠す切り替え', '_このブックの字たち']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★式の セルは 見ない★', /if \(x\.f && String\(x\.f\)\[0\] === '='\) continue;/.test(book));
ok('★変更内容は 履歴の 画面を 開く（同じ物を 2つ 作らない）★',
  /openRireki\('校閲'\)/.test(抜く('変更内容を見る') || ''));
ok('★typeof で 守らない（黙って 素通りしない）★',
  !/typeof openRireki/.test(抜く('変更内容を見る') || ''));
ok('★共有していない事を 実測どおり 言う★', /この ブックは 共有していません/.test(book));
ok('★インクは 消さずに しまうだけ★', /手書きを しまいました（消えては いません）/.test(book));
ok('★隠している 時は 描かない★', /if \(!インクを隠す\) インクを描く\(\);/.test(book));
ok('★日本語は 見ない事を 画面に 書いてある★', /日本語は 見ません/.test(book));
ok('★大きな 辞書を 持たない事を 画面に 書いてある★', /大きな 辞書を 持ちません/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑦ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('スペルチェック', 'スペルを見る');
  試す('類義語辞典', '類義語を聞く');
  試す('見やすさ検査', '見やすさを見る');
  試す('翻訳', '翻訳を聞く');
  試す('変更内容を表示', '変更内容を見る');
  試す('共有を解除', '共有を解除');
  試す('インクを非表示', 'インクを隠す切り替え');
}

console.log('\nreview2: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (R.スペルを見る([{ 値: 'recieve', 場所: 'A1' }]).length === 0) {
    素通り++; console.log('  ★素通り★ recieve を 見逃した');
  } else console.log('  ok   recieve を 見つけた');
  if (R.スペルを見る([{ 値: 'receive', 場所: 'A1' }]).length !== 0) {
    素通り++; console.log('  ★素通り★ 正しい receive を あやしいと 言った');
  } else console.log('  ok   receive は 通した');
  if (R.スペルを見る([{ 値: 'こんにちは', 場所: 'A1' }]).length !== 0) {
    素通り++; console.log('  ★素通り★ 日本語を あやしいと 言った（実Excel は 通す）');
  } else console.log('  ok   日本語は 通した');
  if (Math.abs(R.明るさの比('#000000', '#FFFFFF') - 21) > 0.01) {
    素通り++; console.log('  ★素通り★ 黒と白の 比が 21 で ない');
  } else console.log('  ok   黒と白は 21:1');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
