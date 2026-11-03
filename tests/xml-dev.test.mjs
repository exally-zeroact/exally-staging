/* xml-dev.test.mjs — ★開発タブ（XML の 対応付け・読み書き／コントロール）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-xml.ps1 / tools/measure-xml2.ps1
 *    対応付けの 名前 … ★`people_対応付け`（ルート ＋ `_対応付け`）★／`RootElementName`=`people`
 *    ★`IsExportable` は 結ぶ前 False・セルに 結んだ後 True★
 *    ★結ぶ前に `Import` を 呼ぶと 断られる★
 *      「対応付けられている要素がないためデータはインポートされませんでした。
 *        Range.XPath.SetValue を使用して XML 要素をシートに対応付けてください。」
 *    結んでから `Import` → ★戻り 0（成功）★
 *      入るのは ★A1='ta' B1='20' の 1行だけ★（何行も 入れるには ★表が 要る★）
 *    `Export` … `IsExportable` が False だと ★断られる★
 *      出た XML に ★`standalone="yes"`★ と ★`xmlns:xsi=…XMLSchema-instance`★ が 付く
 *    結んでいない セル（A2）を 直しても ★出る 中身は 変わらない★
 *    決めの 既定 … AdjustColumnWidth=True／★AppendOnImport=False★／
 *      PreserveColumnFilter=True／PreserveNumberFormatting=True／
 *      SaveDataSourceDefinition=True／★ShowImportExportValidationErrors=False★
 *    ボタン … `Button 1`／★90 × 24★／字 `ボタン 1`／★つないだ マクロは 空★
 *    ドロップダウン … `Drop Down 2`／中の 数 2／★えらんでいる 番号は はじめ 0★
 *
 *  走らせ方: node tests/xml-dev.test.mjs [--self-test]
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
const X = require_(path.join(ROOT, 'lib/xml-map.js'));

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

/* ★実Excel に 読ませた のと 同じ XML★ */
const XML = ['<?xml version="1.0" encoding="UTF-8"?>', '<people>',
  '  <person><name>ta</name><age>20</age></person>',
  '  <person><name>ni</name><age>30</age></person>', '</people>'].join('\n');

console.log('\n[① 測った 道具が 残っている]');
ok('tools/measure-xml.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-xml.ps1')));
ok('tools/measure-xml2.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-xml2.ps1')));

console.log('\n[② XML を 読む]');
const 木 = X.木にする(XML);
ok('★ルートは people★', 木 && 木.名 === 'people', 木 && 木.名);
ok('★中に person が 2つ★', 木.子.length === 2 && 木.子[0].名 === 'person');
ok('★対応付けの 名前は people_対応付け（実測）★',
  X.名前を決める(木.名) === 'people_対応付け', X.名前を決める(木.名));
/* ★2つの 読み取り（ブラウザの DOMParser と 手作りの 物）で 答えを 合わせる★
   08-30 実測＝手作りの 方が 閉じ札の 名前違いを 通していた */
ok('★閉じ札の 名前が ちがう XML は null★', X.木にする('<a><b></a>') === null);
ok('★閉じ忘れも null★', X.木にする('<a><b></b>') === null);
ok('★根が 2つも null★', X.木にする('<a></a><b></b>') === null);
ok('  ただの 字も null', X.木にする('これは XML ではない') === null);
ok('  空も null', X.木にする('') === null);
ok('  自己で 閉じる 札は 読める', !!X.木にする('<a><b/></a>'));
ok('  宣言が 付いていても 読める', !!X.木にする('<?xml version="1.0"?><a>1</a>'));
{
  const 道 = X.道を集める(木);
  const 葉 = 道.filter(v => v.葉).map(v => v.道);
  ok('★葉は name と age の 2本★',
    葉.length === 2 && 葉[0] === '/people/person/name' && 葉[1] === '/people/person/age',
    JSON.stringify(葉));
  ok('★person は くり返すと 分かる★',
    道.filter(v => v.道 === '/people/person')[0].くり返す === true);
  ok('  ルートは くり返さない', 道.filter(v => v.道 === '/people')[0].くり返す === false);
  ok('  同じ 道を 2回 出さない', new Set(道.map(v => v.道)).size === 道.length);
}
ok('★値を 全部 取れる（ta と ni）★',
  JSON.stringify(X.値を取る(木, '/people/person/name')) === '["ta","ni"]',
  JSON.stringify(X.値を取る(木, '/people/person/name')));
ok('★age も 2つ★', JSON.stringify(X.値を取る(木, '/people/person/age')) === '["20","30"]');
ok('  無い 道は 空', X.値を取る(木, '/people/nothing').length === 0);
ok('  ルート違いは 空', X.値を取る(木, '/other/person/name').length === 0);

console.log('\n[③ ★結ぶ前は 出せない（実測 IsExportable=False）★]');
ok('★結ぶ前は False★', X.出せるか([]) === false);
ok('★1本 結べば True★', X.出せるか([{ 道: '/people/person/name', r: 0, c: 0 }]) === true);
ok('★結ぶ前に 書き出すと null（実Excel も 断る）★',
  X.書き出す(木, [], () => '') === null);

console.log('\n[④ 読み込み（★うちは くり返す 物を 何行も 入れる★）]');
{
  const 結び = [{ 道: '/people/person/name', r: 0, c: 0 }, { 道: '/people/person/age', r: 0, c: 1 }];
  const 入 = X.読み込む(木, 結び);
  ok('★4個 入る（2行 × 2列）★', 入.length === 4, String(入.length));
  const 表 = {};
  入.forEach(v => { 表[v.r + ',' + v.c] = v.値; });
  ok('★A1=ta（実Excel と 同じ）★', 表['0,0'] === 'ta', 表['0,0']);
  ok('★B1=20（実Excel と 同じ）★', 表['0,1'] === '20', 表['0,1']);
  /* ★ここが 違う所★＝実Excel は 表が 無いと 1行だけ。うちは 2行目も 入れる */
  ok('★A2=ni（★実Excel は 表が 要る／うちは そのまま★）★', 表['1,0'] === 'ni', 表['1,0']);
  ok('★B2=30★', 表['1,1'] === '30', 表['1,1']);
  ok('  結びが 空なら 何も 入らない', X.読み込む(木, []).length === 0);
}

console.log('\n[⑤ 書き出し（実測＝standalone="yes" と xmlns:xsi が 付く）]');
{
  const 結び = [{ 道: '/people/person/name', r: 0, c: 0 }, { 道: '/people/person/age', r: 0, c: 1 }];
  const 表 = { '0,0': 'ta', '0,1': '20', '1,0': 'ni', '1,1': '30' };
  const 字 = X.書き出す(木, 結び, (r, c) => 表[r + ',' + c]);
  ok('★standalone="yes" が 付く（実測）★', /standalone="yes"/.test(字));
  ok('★xmlns:xsi が 付く（実測）★',
    /xmlns:xsi="http:\/\/www\.w3\.org\/2001\/XMLSchema-instance"/.test(字));
  ok('★ルートは people★', /<people /.test(字) && /<\/people>/.test(字));
  ok('★person が 2つ 出る★', (字.match(/<person>/g) || []).length === 2);
  ok('  中身が 入っている', /<name>ta<\/name>/.test(字) && /<age>30<\/age>/.test(字));
  ok('★< > & は 安全に する★',
    /&lt;a&gt;/.test(X.書き出す(木, [結び[0]], () => '<a>')), X.書き出す(木, [結び[0]], () => '<a>'));
  ok('★空の セルで 止まる（下まで 数えない）★',
    (X.書き出す(木, 結び, (r, c) => (r === 0 ? 表[r + ',' + c] : ''))
      .match(/<person>/g) || []).length === 1);
}

console.log('\n[⑥ 対応付けの 決め（実測の 既定）]');
{
  const k = X.既定の決め;
  ok('★列の 幅を 合わせる = する★', k.列の幅を合わせる === true);
  ok('★読み込みで 下に 足す = ★しない★（実測 False）★', k.読み込みで足す === false);
  ok('  絞り込みを 残す = する', k.絞り込みを残す === true);
  ok('  書式を 残す = する', k.書式を残す === true);
  ok('  元の 場所を 覚える = する', k.元の場所を覚える === true);
  ok('★検査の 間違いを 見せる = ★しない★（実測 False）★', k.検査の間違いを見せる === false);
}

console.log('\n[⑦ 画面から 押せる]');
for (const n of ['XMLの窓を開く', 'XMLを読む', 'XMLの道を出す', 'XMLの結びを探す', 'XMLを結ぶ',
  'XMLを読み込む', 'XMLを書き出す', 'XMLの決めを見る', 'XMLを更新', 'XML拡張パック',
  'コントロールを挿入', 'コントロールを置く', 'コントロールのプロパティ',
  'コントロールの決めを決める', 'コントロールを押す', 'コントロールを描く']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★結ぶ前に 読み込むと 断る（実Excel と 同じ 言い方）★',
  /対応付けられている 要素が ありません/.test(book));
ok('★結ぶ前に 出すと 断る（IsExportable に 触れる）★',
  /IsExportable が False の 時は 断ります/.test(book));
ok('★ボタンは 90 × 24（実測）★', /w: 90, h: 24,/.test(抜く('コントロールを置く') || ''));
ok('★ボタンの 字は 「ボタン 1」（実測）★', /'ボタン 1'/.test(book));
ok('★えらんでいる 番号は はじめ 0（実測）★', /選び: 0 \}/.test(book));
/* ★字の 頭の 引用符 から 当てる★（2026-10-09）＝★を 外したら 同じ 字が book.html の 注記にも 在り、
     画面の 字 2か所を 消しても 緑の まま だった（対立役が 試した） */
ok('★VBA を 動かさない事を 画面に 書いてある★', /'うちは VBA を 動かしません/.test(book));
ok('★ボタンの 送り先は AI 1つだけ★', /return AIへ送る\(o\.レシピ\);/.test(抜く('コントロールを押す') || ''));
ok('★typeof で 守っていない★', !/typeof レシピを動かす/.test(book));
ok('★出す 口は FileOut 1つ★', /FileOut\.deliver\(字, /.test(抜く('XMLを書き出す') || ''));
ok('★入り口（XMLファイル）が 在る★', /id="xmlFileInput"/.test(book));
ok('★コントロールは 押せる（描くだけでは ない）★',
  /_o\.種類 === 'コントロール'/.test(book));
ok('★くり返す 物の 違いを 画面に 書いてある★', /くり返す 物を 何行も 入れるのに 表が 要ります/.test(book));
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
  試す('XMLソース', 'XMLの窓を開く');
  試す('XML対応付けの決め', 'XMLの決めを見る');
  試す('XML拡張パック', 'XML拡張パック');
  試す('XMLデータの更新', 'XMLを更新');
  試す('XMLインポート', 'XMLの窓を開く');
  試す('XMLエクスポート', 'XMLを書き出す');
  試す('コントロール挿入', 'コントロールを挿入');
  試す('コントロールの決め', 'コントロールのプロパティ');
}

console.log('\nxml-dev: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (X.出せるか([]) !== false) { 素通り++; console.log('  ★素通り★ 結ぶ前でも 出せると 言う'); }
  else console.log('  ok   結ぶ前は 出せない');
  const 結び = [{ 道: '/people/person/name', r: 0, c: 0 }];
  const 字 = X.書き出す(木, 結び, () => 'ta');
  if (!/standalone="yes"/.test(字)) { 素通り++; console.log('  ★素通り★ standalone が 付いていない'); }
  else console.log('  ok   standalone が 付く');
  if (X.名前を決める('people') !== 'people_対応付け') {
    素通り++; console.log('  ★素通り★ 対応付けの 名前が 実測と 違う');
  } else console.log('  ok   対応付けの 名前が 実測と 同じ');
  if (X.値を取る(木, '/people/person/name').length !== 2) {
    素通り++; console.log('  ★素通り★ くり返す 値を 全部 取れていない');
  } else console.log('  ok   くり返す 値を 2つ 取れる');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
