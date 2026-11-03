/* symbol-form.test.mjs — ★記号と特殊文字★ と ★フォーム★ 2026-08-30
 *
 *  ★実測（実Excel 16.0）★
 *    ・フォーム … ShowDataForm は 在る／相手は ★まわりの 塊（CurrentRegion）★
 *      （A2 を 選ぶと A1:B3＝★見出しも 入る★）… tools/measure-form.ps1
 *    ・記号 … ★窓の 中に 並ぶ 字は COM から 読めない（未測定）★
 *      AutoCorrect の 置き換え表も 0件だった … tools/measure-symbol.ps1
 *      ⇒ 並びは ★うちの決め★。決めた 理由は lib/symbols.js の 見出しに 書いた。
 *
 *  走らせ方: node tests/symbol-form.test.mjs [--self-test]
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

const S = require_(path.join(ROOT, 'lib/symbols.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const lib = fs.readFileSync(path.join(ROOT, 'lib/symbols.js'), 'utf8');

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

console.log('\n[① ★測れなかった事を 測れたと 書いていない★]');
ok('測る道具が 残っている（記号）', fs.existsSync(path.join(ROOT, 'tools/measure-symbol.ps1')));
ok('測る道具が 残っている（フォーム）', fs.existsSync(path.join(ROOT, 'tools/measure-form.ps1')));
ok('★記号の 並びは 未測定と 書いてある★', /★未測定★/.test(lib), lib.slice(0, 200));
ok('★なぜ その並びに したかの 理由が 在る★', /決め方の 理由/.test(lib));
ok('★フォームは 実測（CurrentRegion）と 書いてある★',
  /まわりの 塊（CurrentRegion）/.test(book));

console.log('\n[② 記号の 並び]');
{
  ok('組が 1つ以上 在る', S.記号.length >= 5, String(S.記号.length));
  let 数 = 0;
  const 見た = {};
  let 重なり = 0;
  for (const 組 of S.記号) {
    ok('  ' + 組.組 + ' に 字が 在る', 組.字.length > 0, String(組.字.length));
    for (const c of 組.字) { 数++; if (見た[c]) 重なり++; 見た[c] = true; }
  }
  ok('★同じ字を 2回 出していない★', 重なり === 0, String(重なり) + '個 重なり');
  ok('全部で 60字 以上', 数 >= 60, String(数));
  ok('特殊文字が 10個 以上', S.特殊文字.length >= 10, String(S.特殊文字.length));
  /* ★絵文字は 入れない（紙で 化ける）★ */
  let 絵文字 = 0;
  for (const 組 of S.記号) for (const c of 組.字) if (c.codePointAt(0) > 0xFFFF) 絵文字++;
  for (const t of S.特殊文字) if (t.字.codePointAt(0) > 0xFFFF) 絵文字++;
  ok('★絵文字（上の面の字）を 入れていない★', 絵文字 === 0, String(絵文字));
  /* 名前は 全部 付いている */
  ok('特殊文字は 全部 名前つき', S.特殊文字.every((t) => t.名 && t.字), JSON.stringify(S.特殊文字[0]));
}

console.log('\n[③ 文字コードから 字を 作る]');
for (const [入れる, 期待, なぜ] of [
  ['U+00A9', '©', ''],
  ['u+00a9', '©', '小文字でも'],
  ['00A9', '©', '★U+ が 無くても★'],
  ['0x2122', '™', '0x でも'],
  ['169', '©', '★10進の 169★'],
  ['  U+2026  ', '…', '前後に 空白が 在っても'],
  ['U+1F600', '😀', '★上の面の字も 1文字で★'],
]) {
  const 出 = S.コードから字(入れる);
  ok('「' + 入れる.trim() + '」→ ' + 期待 + (なぜ ? '  ' + なぜ : ''), 出 === 期待, JSON.stringify(出));
}
for (const [入れる, なぜ] of [
  ['', '空'], ['あ', '字'], ['U+', '数が 無い'],
  ['U+D800', '★片割れは 字に ならない★'], ['U+110000', '★上限より 上★'],
  ['-5', 'マイナス'],
]) {
  ok('★読めない物は null★ … 「' + 入れる + '」  ' + なぜ, S.コードから字(入れる) === null,
    JSON.stringify(S.コードから字(入れる)));
}
ok('字 → U+XXXX', S.字からコード('©') === 'U+00A9', S.字からコード('©'));
ok('4桁に 満たない時も 0で 埋める', S.字からコード('A') === 'U+0041', S.字からコード('A'));

console.log('\n[④ 画面に つながっている（記号）]');
for (const n of ['記号の窓を開く', '記号の窓を閉じる', '記号の面を切る', '記号のボタン',
  '記号を入れる', '記号のコードを見る', '記号のコードを入れる']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="symOverlay"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/symbols\.js/.test(book));
ok('★押した字は 後ろに 足す（今の字を 消さない）★', /setCell\(selR1, selC1, 今 \+ 字\)/.test(book));
ok('★読めない コードは 理由を 出す★', /その 文字コードは 読めません/.test(book));

console.log('\n[⑤ 画面に つながっている（フォーム）]');
for (const n of ['フォームの窓を開く', 'フォームを描く', 'フォームを書き込む', 'フォームを動かす',
  'フォームで足す', 'フォームで消す', 'フォームを閉じる']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="formOverlay"/.test(book));
ok('★表が 在れば 表を 相手に する★', /var t = 今いる表\(\);/.test(抜く('フォームの窓を開く')));
ok('★表が 無ければ まわりの 塊★（実測と 同じ 見立て）',
  /表のまわりを当てる\(selR1, selC1\)/.test(抜く('フォームの窓を開く')));
ok('★中身の 行が 無ければ 開かない★', /中身の 行が ありません/.test(book));
ok('★最後の1行は 消せない（空の表に しない）★', /最後の 1行は 消せません/.test(book));
ok('★端まで 行ったら 理由を 出す★', /これが 最後の 行です/.test(book) && /これが 最初の 行です/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ フォームの 中身（本当に 動くか）]');
{
  /* 書き込みの 中身だけ 取り出して 走らせる */
  const f = new Function('sheets', 'activeSheet', '_フォーム', 'document', 'setCell',
    抜く('フォームを書き込む') + '\nreturn フォームを書き込む;');
  const データ = { '0,0': { v: '月', f: '月' }, '1,0': { v: '1月', f: '1月' }, '1,1': { v: 100, f: '100' } };
  const 入れた = [];
  const 箱 = [
    { getAttribute: () => '0', value: '2月' },
    { getAttribute: () => '1', value: '100' },
  ];
  const n = f([{ data: データ }], 0, { r1: 0, c1: 0, r2: 1, c2: 1, 見出し: ['月', '売上'], 今の行: 1 },
    { querySelectorAll: () => 箱 },
    (r, c, v) => { 入れた.push(r + ',' + c + '=' + v); })();
  ok('★変えた所だけ 入れる★', n === 1, String(n));
  ok('  入れた所', JSON.stringify(入れた) === '["1,0=2月"]', JSON.stringify(入れた));
}

console.log('\n[⑦ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [['記号と特殊文字', '記号の窓を開く'], ['フォーム', 'フォームの窓を開く']]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nsymbol-form: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 片割れ（サロゲート）を 通したら 字が 壊れる */
  if (S.コードから字('U+D800') !== null) { 素通り++; console.log('  ★素通り★ 片割れを 通した'); }
  else console.log('  ok   片割れは 通さない');
  /* 壊し② 上の面の字が 2文字に 割れていないか */
  const 顔 = S.コードから字('U+1F600');
  if (!顔 || [...顔].length !== 1) { 素通り++; console.log('  ★素通り★ 上の面の字が 1文字に なっていない'); }
  else console.log('  ok   上の面の字も 1文字');
  /* 壊し③ 記号の 並びが 空に なっていないか */
  if (!S.記号.length || !S.特殊文字.length) { 素通り++; console.log('  ★素通り★ 並びが 空'); }
  else console.log('  ok   並びは 空でない');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
