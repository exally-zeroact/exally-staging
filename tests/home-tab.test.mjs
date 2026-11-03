/* home-tab.test.mjs — ★ホーム（フォント・文字の向き・クリップボード・表として書式設定）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-home-page.ps1
 *    ・既定の フォント ＝ ★游ゴシック／大きさ 11★
 *    ・文字の 向き（Orientation）… 0 → ★-4128★／45 → 45／-45 → -45／
 *      90 → ★-4171★／-90 → ★-4170★／縦書き ★-4166★
 *
 *  ★Office クリップボードは 24個まで★（実Excelの 決め）。
 *
 *  走らせ方: node tests/home-tab.test.mjs [--self-test]
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

console.log('\n[① フォント（実測＝既定は 游ゴシック 11）]');
for (const n of ['入っているフォント', '使えるフォント', 'フォントの窓を開く', 'フォントを決める']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="fontOverlay"/.test(book));
ok('★実Excelの 既定（游ゴシック）を 一番 先に 出す★', /var フォントの候補 = \['游ゴシック'/.test(book));
ok('★機械に 入っているか 本当に 測る（幅を くらべる）★',
  /measureText\(見本\)\.width/.test(book) && /32px monospace/.test(book));
ok('  予備の 字でも もう一度 くらべる（1回だけでは 取りこぼす）', /32px serif/.test(book));
ok('★入っていない フォントは 出さないと 書いてある★', /機械に 入っていない フォントは 出しません/.test(book));
/* ★★2026-09-11 … 字面から 中身の 検査に 変えました★★
   ★変えた 訳★＝字体を 作る 道が ★4か所★に 散って いたので `_マスの書体` 1本に まとめた
     （ふつうの マス／結合した マス／リボンの札／Ctrl+` の 式表示）
   ★守る 中身は 同じ★＝★マスが 自分の 字体を 持って いれば それが 勝つ★
   ★字面だけ 見て いると、中身が 同じでも 赤に なり、
     中身が 変わっても 字面が 残って いれば 緑に なります★ */
{
  const 頭 = book.indexOf('function _マスの書体');
  ok('★字体を 作る 道が 1本に 在る★', 頭 >= 0);
  const 中 = 頭 >= 0 ? book.slice(頭, 頭 + 600) : '';
  ok('★描く時に そのフォントを 使う（マスの 字体が 勝つ）★',
    /if \(cell && cell\.fontName\) return/.test(中) && /Noto Sans JP/.test(中), 中.slice(0, 120));
  /* ★古い 書き方が 残って いないか★＝残ると そこだけ 別の 字に なる */
  ok('★直に 書いて いる 所が 残って いない★',
    (book.match(/cell\.fontName \? \('"'/g) || []).length === 0);
}
ok('★1つも 無い時の 言い方が 在る★', /出せる フォントが 見つかりませんでした/.test(book));

console.log('\n[② 文字の 向き（実測の 数と 同じ）]');
for (const n of ['文字の向きの窓を開く', '文字の向きを決める']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="orientOverlay"/.test(book));
{
  const 並び = book.slice(book.indexOf('var 向きの並び = ['), book.indexOf('function 文字の向きの窓を開く'));
  for (const [名, 数] of [['横', '-4128'], ['上へ', '-4171'], ['下へ', '-4170'], ['縦書き', '-4166']]) {
    ok('★' + 名 + ' の 数は ' + 数 + '（実測）★', 並び.indexOf('Excel: ' + 数) >= 0, 並び.slice(0, 60));
  }
  ok('45 と -45 も 在る', /角: 45, Excel: 45/.test(並び) && /角: -45, Excel: -45/.test(並び));
  ok('★6通り 全部★', (並び.match(/Excel:/g) || []).length === 6, String((並び.match(/Excel:/g) || []).length));
}
ok('★描く時に 回している★', /ctx\.rotate\(-_向き \* Math\.PI \/ 180\)/.test(book));
ok('★縦書きは 1文字ずつ 下へ★', /縦書き＝1文字ずつ 下へ 並べる/.test(book));
ok('★横（0）に 戻したら 印を 消す★', /if \(値 === 0\) delete cell\.向き;/.test(book));

console.log('\n[③ ★新しい 書式を 置いていかない★]');
{
  const fmt = book.slice(book.indexOf('var FMT = ['), book.indexOf('var FMT = [') + 300);
  ok('引っぱる時の 一覧に fontName が 在る', fmt.indexOf("'fontName'") >= 0, fmt);
  ok('引っぱる時の 一覧に 向き が 在る', fmt.indexOf("'向き'") >= 0, fmt);
  const 書 = book.slice(book.indexOf('var 書式のキー = ['), book.indexOf('function 中身を消す'));
  ok('書式を 消す/写す 一覧に fontName が 在る', 書.indexOf("'fontName'") >= 0, 書);
  ok('書式を 消す/写す 一覧に 向き が 在る', 書.indexOf("'向き'") >= 0, 書);
}

console.log('\n[④ クリップボード（24個まで）]');
for (const n of ['クリップに足す', 'クリップボードを開く', 'クリップボードを描く',
  'クリップから貼る', 'クリップから消す', 'クリップボードを空に']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る', /id="clipOverlay"/.test(book));
ok('★24個まで（実Excelと 同じ）★', /var クリップの上限 = 24/.test(book));
ok('★コピーしたら 覚える★', /クリップに足す\(copyBuffer\)/.test(book));
{
  const f = new Function('クリップの箱', 'クリップの上限',
    抜く('クリップに足す') + '\nreturn function(x){ return クリップに足す(x); };');
  const 箱 = [];
  const 足す = f(箱, 3);
  足す({ text: 'あ' }); 足す({ text: 'い' }); 足す({ text: 'う' }); 足す({ text: 'え' });
  ok('★上限を 超えたら 古い物から 消える★', 箱.length === 3, String(箱.length));
  ok('★新しい物が 先頭★', 箱[0].text === 'え', JSON.stringify(箱.map((x) => x.text)));
  const 前 = 箱.length;
  足す(null); 足す({ text: '' });
  ok('★空は 覚えない★', 箱.length === 前, String(箱.length));
}
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑤ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['フォント', 'フォントの窓を開く'], ['方向', '文字の向きの窓を開く'],
    ['テーブルとして書式設定', '表の窓を開く'], ['Officeクリップボード', 'クリップボードを開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nhome-tab: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 向きの 数が 実測と 違っていないか */
  const 並び = book.slice(book.indexOf('var 向きの並び = ['), book.indexOf('function 文字の向きの窓を開く'));
  if (並び.indexOf('Excel: -4171') < 0 || 並び.indexOf('Excel: -4170') < 0) {
    素通り++; console.log('  ★素通り★ 90 / -90 の 数が 実測（-4171 / -4170）と 違う');
  } else console.log('  ok   90 → -4171 ／ -90 → -4170');
  /* 壊し② 書式の 一覧から 落ちていないか */
  if (book.indexOf("'fontName','向き'") < 0 && book.indexOf("'fontName', '向き'") < 0) {
    素通り++; console.log('  ★素通り★ 書式の 一覧に 新しい物が 入っていない');
  } else console.log('  ok   書式の 一覧に 入っている');
  /* 壊し③ フォントが 在るかを 測らずに 全部 出していないか */
  if (!/measureText\(見本\)\.width/.test(book)) {
    素通り++; console.log('  ★素通り★ 入っているか 測っていない');
  } else console.log('  ok   入っているか 測っている');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
