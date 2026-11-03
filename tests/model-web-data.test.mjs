/* model-web-data.test.mjs — ★3Dモデル／Webから／データモデル／データ分析★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-xml.ps1 と 追加の 実測
 *    ・表（ListObject）は 名前を 付けられる／`ListColumns` で 列の 名が 取れる
 *    ・★表の 既定の 名前は「テーブル3」★（日本語＋番号）
 *    ・`Workbook.Model.ModelTables.Count` … ★はじめ 0★
 *      （＝★表を 作っただけでは モデルに 入らない★）
 *    ・`Model.ModelRelationships.Count` … ★0★
 *    ・`Model.AddConnection` … ★COM から 呼べなかった★
 *    ・3D モデル … ★形（Shapes）の 1つ★／★中の 三角形の 数は 読めなかった★
 *    ・データ分析 … ★Microsoft の クラウドで 動く（機械からは 呼べない）★
 *
 *  ★真似る 相手の 数字が 無い 所は 自分で 決めて 自分で 測る★
 *
 *  走らせ方: node tests/model-web-data.test.mjs [--self-test]
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
const M3 = require_(path.join(ROOT, 'lib/model3d.js'));
const DM = require_(path.join(ROOT, 'lib/data-model.js'));

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

console.log('\n[① 3D モデル（.obj を 読む）]');
{
  const c = M3.立方体();
  ok('★立方体の 点は 8個★', c.点.length === 8, String(c.点.length));
  ok('★四角 6面 → 三角 12面に 切る★', c.面.length === 12, String(c.面.length));
  ok('★同じ 辺を 2回 出さない★', M3.辺たち(c).length === 18, String(M3.辺たち(c).length));
  ok('  箱が 合う', (() => {
    const b = M3.箱(c);
    return b.x1 === -1 && b.x2 === 1 && b.y1 === -1 && b.y2 === 1 && b.z1 === -1 && b.z2 === 1;
  })());
  /* ★どんな 向きに 回しても 箱から はみ出さない★（対角線で 合わせている） */
  let 全部 = true, 最悪 = 0;
  for (let a = 0; a < Math.PI * 2; a += 0.2) {
    for (let b = -1.5; b <= 1.5; b += 0.3) {
      for (const q of M3.画面の点(c, a, b, 200, 200)) {
        if (q.x < 0 || q.x > 200 || q.y < 0 || q.y > 200) 全部 = false;
        最悪 = Math.max(最悪, Math.abs(q.x - 100), Math.abs(q.y - 100));
      }
    }
  }
  ok('★どの 向きでも はみ出さない（' + Math.round(最悪) + '/100）★', 全部, String(最悪.toFixed(1)));
  ok('★回すと 絵が 変わる★',
    JSON.stringify(M3.画面の点(c, 0, 0, 200, 200)) !== JSON.stringify(M3.画面の点(c, 1, 0.5, 200, 200)));
  ok('★とても 大きい モデルでも 入る★', (() => {
    const 大 = M3.objを読む('v -1000 -1000 -1000\nv 1000 1000 1000\nv 0 500 0\nf 1 2 3');
    return M3.画面の点(大, 1, 0.5, 100, 100).every(q => q.x >= 0 && q.x <= 100 && q.y >= 0 && q.y <= 100);
  })());
  console.log('  ─ .obj の 読み方 ─');
  ok('  空は null', M3.objを読む('') === null && M3.objを読む('   ') === null);
  ok('  点が 無ければ null', M3.objを読む('f 1 2 3') === null);
  ok('  `#` の 行は 読み飛ばす', M3.objを読む('# めも\nv 0 0 0').点.length === 1);
  ok('★`1/2/3` の 形も 読める★',
    JSON.stringify(M3.objを読む('v 0 0 0\nv 1 0 0\nv 0 1 0\nf 1/1/1 2/2/2 3/3/3').面) === '[[0,1,2]]');
  ok('★`1//3` の 形も 読める★',
    JSON.stringify(M3.objを読む('v 0 0 0\nv 1 0 0\nv 0 1 0\nf 1//1 2//2 3//3').面) === '[[0,1,2]]');
  ok('★負の 番号は 後ろから★',
    JSON.stringify(M3.objを読む('v 0 0 0\nv 1 0 0\nv 0 1 0\nf -3 -2 -1').面) === '[[0,1,2]]');
  ok('  数で ない 点は 飛ばす', M3.objを読む('v a b c\nv 0 0 0').点.length === 1);
  ok('  5角形は 三角 3つに なる',
    M3.objを読む('v 0 0 0\nv 1 0 0\nv 2 1 0\nv 1 2 0\nv 0 2 0\nf 1 2 3 4 5').面.length === 3);
}

console.log('\n[② データ モデル（表どうしを つなげる）]');
{
  const 人 = { 名: '人', 行たち: [['id', '名'], [1, 'あ'], [2, 'い'], [3, 'う']] };
  const 売 = { 名: '売上', 行たち: [['id', '金額'], [1, 100], [2, 200]] };
  const 重 = { 名: '重', 行たち: [['id', 'x'], [1, 'a'], [1, 'b']] };
  ok('★1つずつか を 見分ける★', DM.一つずつか(人, 0) === true && DM.一つずつか(重, 0) === false);
  ok('  空の セルは 数えない', DM.一つずつか({ 行たち: [['id'], [''], ['']] }, 0) === true);
  ok('  列の 番号を 引ける', DM.列の番号(人, '名') === 1 && DM.列の番号(人, 'ない') === -1);
  ok('★つなげられる（売→人）★', DM.つなげるか(売, 'id', 人, 'id').出来る === true);
  ok('★1つずつで ない 側は 断る★', (() => {
    const r = DM.つなげるか(人, 'id', 重, 'id');
    return r.出来る === false && /同じ 値が 2回/.test(r.訳);
  })());
  ok('  無い 列は 断る', DM.つなげるか(売, 'ない', 人, 'id').出来る === false);
  ok('★合う 値が 1つも 無ければ 断る★', (() => {
    const r = DM.つなげるか({ 名: 'x', 行たち: [['id'], [99]] }, 'id', 人, 'id');
    return r.出来る === false && /合う 値が 1つも/.test(r.訳);
  })());
  const 出 = DM.つなげる(人, 'id', 売, 'id');
  ok('★つなげた 表は 見出し ＋ 3行★', 出.length === 4, String(出.length));
  ok('★列の 名は 「表.列」★', 出[0][2] === '売上.金額', 出[0][2]);
  ok('★同じ 列を 2回 出さない★', 出[0].length === 3, JSON.stringify(出[0]));
  ok('  値が 入る', 出[1][2] === 100 && 出[2][2] === 200);
  ok('★合わない 行（id=3）も 残る★', 出[3][0] === 3 && 出[3][2] === '', JSON.stringify(出[3]));
  ok('  つなげられない 時は null', DM.つなげる(人, 'id', 重, 'id') === null);
  const t = [];
  DM.つなぎを足す(t, '人', 'id', '売上', 'id');
  DM.つなぎを足す(t, '人', 'id', '売上', 'id');
  ok('★同じ つなぎは 1本だけ★', t.length === 1, String(t.length));
  ok('  消せる', DM.つなぎを消す(t, 0) === true && t.length === 0);
  ok('  無い 番号は false', DM.つなぎを消す(t, 5) === false);
}

console.log('\n[③ 画面から 押せる]');
for (const n of ['三次元を開く', '三次元を読む', '三次元を置く', '三次元を回す', '三次元を描く',
  '三次元の窓を開く', 'Webから読む', '_表をHTMLから', 'Webの字から入れる', '_表を入れる',
  'Webを取りに行く', '_この画面の表たち', 'データモデルを開く', 'つながりを足す',
  'つながりを消す', 'つなげて出す', 'データ分析']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('★.obj だけ 読めると 言う★', /\.obj の ファイルだけ 読めます/.test(book));
ok('★.glb / .fbx は 読めないと 画面に 書いてある★',
  /\.glb や \.fbx は 中身の 決まりが 大きくて 読めません/.test(book));
ok('★線だけだと 画面に 書いてある★', /線だけで 出します/.test(book));
ok('★はみ出さないと 画面に 書いてある★', /どんな 向きに 回しても 箱から はみ出しません/.test(book));
ok('★向こうが 許した 時だけ 読めると 画面に 書いてある★',
  /向こうの ページが 許した 時だけ 読めます/.test(book));
ok('★読めなかった 訳を 隠さない★', /読めませんでした（/.test(book));
ok('★https から 始まらないと 断る★', /https:\/\/ から 始まる 所を 書いてください/.test(book));
ok('★表が 無ければ 断る★', /表（table）が 見つかりません/.test(book));
ok('★1つずつで ないと つなげない事を 画面に 書いてある★',
  /つながれる 側は 1つずつ（同じ 値が 2回 出ない）で ないと いけません/.test(book));
ok('★ModelTables は はじめ 0 という 実測を 画面に 書いてある★',
  /実Excel は 表を 作っただけでは モデルに 入りません（ModelTables は 0 でした）/.test(book));
ok('★同じ 表どうしは つなげない★', /同じ 表どうしは つなげません/.test(book));
ok('★データ分析は AI に 聞く（お金の 数は うちの 計算）★',
  /合計や 割り算の 数は うちが 計算した 物を 使ってください/.test(book));
ok('★typeof で 守っていない（つなぎたち）★', !/typeof つなぎたち/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[④ 副題を 決めていない 窓が 増えていないか]');
{
  const 行 = book.split(String.fromCharCode(10));
  const 抜け = [];
  for (let i = 0; i < 行.length; i++) {
    if (!/getElementById\('funcTitle'\)\.textContent =/.test(行[i])) continue;
    if (!/窓の副題\(/.test(行.slice(i, i + 5).join(String.fromCharCode(10)))) 抜け.push(i + 1);
  }
  ok('★副題を 決めていない 窓は 0個★', 抜け.length === 0, 抜け.join(' / '));
}

console.log('\n[⑤ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  const 試す = (ボタン, 呼ぶ名) => {
    let 受け = 'よばれていない';
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w; ACT[ボタン](); g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  };
  試す('三次元モデル', '三次元の窓を開く');
  試す('Webから', 'Webから読む');
  試す('データモデル', 'データモデルを開く');
  試す('データ分析', 'データ分析');
}

console.log('\nmodel-web-data: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const c = M3.立方体();
  let はみ出し = false;
  for (const q of M3.画面の点(c, 1, 0.5, 200, 200)) {
    if (q.x < 0 || q.x > 200 || q.y < 0 || q.y > 200) はみ出し = true;
  }
  if (はみ出し) { 素通り++; console.log('  ★素通り★ 回すと はみ出す'); }
  else console.log('  ok   回しても はみ出さない');
  const 人 = { 名: '人', 行たち: [['id'], [1], [2]] };
  const 重 = { 名: '重', 行たち: [['id'], [1], [1]] };
  if (DM.つなげるか(人, 'id', 重, 'id').出来る) { 素通り++; console.log('  ★素通り★ 1つずつで ない 側に つなげてしまう'); }
  else console.log('  ok   1つずつで ない 側は 断る');
  if (M3.objを読む('f 1 2 3') !== null) { 素通り++; console.log('  ★素通り★ 点が 無い のに 読めたと 言う'); }
  else console.log('  ok   点が 無ければ null');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
