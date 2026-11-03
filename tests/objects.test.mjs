/* objects.test.mjs — ★シートに 浮かぶ 物（画像・図形・テキスト）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-shapes.ps1 / tools/measure-colors.ps1
 *    ・図形（四角）… 名前 'Rectangle 1'／塗り ★#156082★／線 ★#042433★／線の太さ ★1.5★
 *    ・テキスト ボックス … 名前 'TextBox 2'／塗り ★#FFFFFF★（見える）／線も 見える
 *    ・重なりの 順（ZOrderPosition）… ★1 から★／前面へ で 2・背面へ で 1
 *    ・図形の 種類の 番号 … 四角=1・角丸=5・円=9・三角=7・右矢印=33・星5=92
 *    ・図形にも ★字が 入る★（TextFrame2 在り）
 *
 *  ★色は 写さない★（司さん「訴えられんような 見せ方で 同じように」）
 *    ＝形・働き・重なりの 決まりは 同じ／色は うちの 緑。実測値は 部品の 見出しに 残す。
 *
 *  走らせ方: node tests/objects.test.mjs [--self-test]
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
const O = require_(path.join(ROOT, 'lib/objects.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const lib = fs.readFileSync(path.join(ROOT, 'lib/objects.js'), 'utf8');

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

console.log('\n[① 測った 値を 書き残してある]');
ok('tools/measure-shapes.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-shapes.ps1')));
ok('tools/measure-colors.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-colors.ps1')));
ok('★塗り #156082・線 #042433 を 書き残した★', /#156082/.test(lib) && /#042433/.test(lib));
ok('★線の 太さ 1.5 は 同じに した★', O.線の太さ === 1.5, String(O.線の太さ));
ok('★色は 写さないと 書いてある★', /色は 写さない/.test(lib));
ok('★うちの 色は Excelの 実測値では ない★',
  O.色.塗り !== '#156082' && O.色.線 !== '#042433', JSON.stringify(O.色));

console.log('\n[② 形の 番号（実測）]');
{
  const 見 = {};
  for (const f of O.形たち) 見[f.名] = f.Excel;
  ok('四角 = 1', 見['四角'] === 1);
  ok('角丸四角 = 5', 見['角丸四角'] === 5);
  ok('丸 = 9', 見['丸'] === 9);
  ok('三角 = 7', 見['三角'] === 7);
  ok('右矢印 = 33', 見['右矢印'] === 33);
  ok('星 = 92', 見['星'] === 92);
  ok('線は 番号なし（AddLine）', 見['線'] === null);
}

console.log('\n[③ 重なりの 順（実測＝1 から／前面 2・背面 1）]');
{
  const 箱 = [{ 名: 'A', z: 1 }, { 名: 'B', z: 2 }];
  O.前面へ(箱, 箱[0]);
  ok('★A を 前面へ → A=2 B=1★', 箱[0].z === 2 && 箱[1].z === 1, JSON.stringify(箱));
  O.背面へ(箱, 箱[0]);
  ok('★A を 背面へ → A=1 B=2★', 箱[0].z === 1 && 箱[1].z === 2, JSON.stringify(箱));
  const 箱2 = [{ z: 5 }, { z: 9 }, { z: 7 }];
  O.順をそろえる(箱2);
  ok('★1 から 続き番号に そろえる★',
    箱2.map((x) => x.z).sort().join(',') === '1,2,3', JSON.stringify(箱2.map((x) => x.z)));
}

console.log('\n[④ どれを 押したか（上に 在る 物が 勝つ）]');
{
  const 箱 = [
    { 名: '下', x: 0, y: 0, w: 100, h: 100, z: 1 },
    { 名: '上', x: 50, y: 50, w: 100, h: 100, z: 2 },
  ];
  ok('★重なった 所は 上が 勝つ★', O.どれを押したか(箱, 60, 60).名 === '上',
    String(O.どれを押したか(箱, 60, 60).名));
  ok('下だけの 所は 下', O.どれを押したか(箱, 10, 10).名 === '下');
  ok('何も 無い所は null', O.どれを押したか(箱, 300, 300) === null);
  ok('空でも 落ちない', O.どれを押したか([], 1, 1) === null);
}

console.log('\n[⑤ 名前（実Excel は 種類＋番号）]');
ok('1つ目は 四角 1', O.名前を決める([], '四角') === '四角 1', O.名前を決める([], '四角'));
ok('2つ目は 四角 2', O.名前を決める(['四角 1'], '四角') === '四角 2');
ok('★空いている 番号を 使う★', O.名前を決める(['四角 1', '四角 3'], '四角') === '四角 2');

console.log('\n[⑥ 描く（形ごとに 違う絵か）]');
{
  const 数える = (種類) => {
    const 出 = { fillRect: 0, stroke: 0, fill: 0, ellipse: 0, moveTo: 0, lineTo: 0, quad: 0, strokeRect: 0 };
    const 何も = () => {};
    const c = {
      save: 何も, restore: 何も, beginPath: 何も, closePath: 何も, setLineDash: 何も,
      moveTo: () => 出.moveTo++, lineTo: () => 出.lineTo++,
      quadraticCurveTo: () => 出.quad++, ellipse: () => 出.ellipse++,
      fillRect: () => 出.fillRect++, strokeRect: () => 出.strokeRect++,
      stroke: () => 出.stroke++, fill: () => 出.fill++, fillText: 何も, drawImage: 何も,
      fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: '', textBaseline: '',
    };
    O.描く(c, { 種類: 種類, x: 0, y: 0, w: 100, h: 50 }, false);
    return JSON.stringify(出);
  };
  const 絵 = {};
  for (const k of ['四角', '角丸四角', '丸', '三角', '右矢印', '星', '線', 'テキスト']) 絵[k] = 数える(k);
  ok('★丸は 楕円を 使う★', JSON.parse(絵['丸']).ellipse === 1, 絵['丸']);
  ok('★角丸は 曲線を 使う★', JSON.parse(絵['角丸四角']).quad === 4, 絵['角丸四角']);
  ok('★星は 10個の 角★', JSON.parse(絵['星']).lineTo === 9, 絵['星']);
  ok('★線は 塗らない★', JSON.parse(絵['線']).fill === 0, 絵['線']);
  ok('四角と 三角は 別の絵', 絵['四角'] !== 絵['三角'], 絵['四角'] + ' / ' + 絵['三角']);
  ok('四角と 右矢印は 別の絵', 絵['四角'] !== 絵['右矢印']);
  ok('テキストは 白で 塗る', JSON.parse(絵['テキスト']).fillRect === 1, 絵['テキスト']);
  ok('★大きさ 0 は 描かない★',
    O.描く({}, { 種類: '四角', x: 0, y: 0, w: 0, h: 0 }, false) === false);
}

console.log('\n[⑦ 画面に つながっている]');
for (const n of ['物の箱', '物を足す', '図形の窓を開く', 'テキストを置く', '画像の窓を開く',
  '画像を読む', '物を描く', '物を押したか', '前面へ移動', '背面へ移動', '物を消す',
  '物の一覧を開く', '物の一覧を描く', '物を選ぶ']) ok(n + ' が 在る', !!抜く(n));
ok('窓が 在る（図形）', /id="shapeOverlay"/.test(book));
ok('板が 在る（一覧）', /id="objListPanel"/.test(book));
ok('入り口が 在る（画像）', /id="imgFileInput"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/objects\.js/.test(book));
ok('★セルの 上に 描く★', /物を描く\(\);/.test(book));
ok('★押した 時は セルより 先に 見る★',
  book.indexOf('var _o = 物を押したか(p.x, p.y);') < book.indexOf('var c=xToC(p.x),r=yToR(p.y);'));
ok('★動かせる／大きさも 変えられる★',
  /touchState\.type==='objmove'/.test(book) && /touchState\.type==='objsize'/.test(book));
ok('★選んでいないのに 前面へ を 押したら 断る★', /先に 物を 押して 選んでください/.test(book));
ok('★画像は 元の 形の まま（つぶさない）★', /元の 形の まま/.test(book));
ok('★読めない 画像は そう 言う★', /その 画像は 読めませんでした/.test(book));
ok('★一覧は 上に 在る 物が 先★', /上に 在る 物が 先/.test(book));
/* ★スクロールしても セルと 一緒に 動く★（画面の 座標で 持たない） */
ok('★物は シートの 座標で 持つ★', /物は「シートの 座標」で 持つ/.test(book));
ok('  画面に 直す 道が 在る', /function 物の画面X\(o\)/.test(book) && /function 物の画面Y\(o\)/.test(book));
ok('  画面から 直す 道も 在る', /function 画面から物のX\(px\)/.test(book));
/* 08-30：写す 時に 持ち物を 数え上げると 新しい 持ち物（中身・回転）が 黙って 落ちたので
   丸ごと 写して x/y だけ 差し替える 形に 変えた（tests/arrange.test.mjs に 見張り） */
ok('★描く時に 画面の 座標へ 直している★', /写し\.x = 物の画面X\(o\); 写し\.y = 物の画面Y\(o\);/.test(book));
ok('★押した所も 直してから 見る★', /どれを押したか\(箱, 画面から物のX\(px\), 画面から物のY\(py\)\)/.test(book));
ok('★動かす時も 直している★', /_sx = 画面から物のX\(p\.x\), _sy = 画面から物のY\(p\.y\)/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑧ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['画像', '画像の窓を開く'], ['図形', '図形の窓を開く'], ['テキスト', 'テキストを置く'],
    ['前面へ移動', '前面へ移動'], ['背面へ移動', '背面へ移動'],
    ['オブジェクトの選択と表示', '物の一覧を開く'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nobjects: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  if (O.色.塗り === '#156082') { 素通り++; console.log('  ★素通り★ Excelの 色を 写している'); }
  else console.log('  ok   Excelの 色を 写していない');
  const 箱 = [{ z: 1 }, { z: 2 }];
  O.前面へ(箱, 箱[0]);
  if (箱[0].z !== 2) { 素通り++; console.log('  ★素通り★ 前面へ が 効いていない … ' + JSON.stringify(箱)); }
  else console.log('  ok   前面へ で 2 に なる');
  const 上下 = O.どれを押したか([{ 名: '下', x: 0, y: 0, w: 9, h: 9, z: 1 },
    { 名: '上', x: 0, y: 0, w: 9, h: 9, z: 2 }], 5, 5);
  if (!上下 || 上下.名 !== '上') { 素通り++; console.log('  ★素通り★ 重なりで 上が 勝っていない'); }
  else console.log('  ok   上が 勝つ');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
