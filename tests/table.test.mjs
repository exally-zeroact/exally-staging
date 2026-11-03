/* table.test.mjs — ★テーブル（Ctrl+T）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-table.ps1 / tools/measure-table2.ps1
 *    名前=テーブル1（2つ目は テーブル2）／見出し行=True／集計行=False／自動フィルタ=True
 *    しま（行）=True・しま（列）=False・最初/最後の列=False
 *    ★1行目が 字 → 見出しに 使う（範囲は そのまま）★
 *    ★1行目も 数 → 見出し行を 1行 足す（D1:E2 → D1:E3）・列の名は 列1・列2★
 *    ★集計行＝左端 '集計'・一番 右の列だけ SUBTOTAL(109,[列名])・間は 空★
 *    ★SUBTOTAL 109 は 隠した行を 数えない★（2行目を 隠して 109=4／9=6）
 *
 *  走らせ方: node tests/table.test.mjs [--self-test]
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

const T = require_(path.join(ROOT, 'lib/table.js'));
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

console.log('\n[① 実測の 道具が 残っている]');
ok('tools/measure-table.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-table.ps1')));
ok('tools/measure-table2.ps1 が 在る', fs.existsSync(path.join(ROOT, 'tools/measure-table2.ps1')));

console.log('\n[② 名前＝テーブル1・テーブル2（実測）]');
ok('1つ目は テーブル1', T.名前を決める([]) === 'テーブル1', T.名前を決める([]));
ok('2つ目は テーブル2', T.名前を決める(['テーブル1']) === 'テーブル2', T.名前を決める(['テーブル1']));
ok('★空いている 番号を 使う★', T.名前を決める(['テーブル1', 'テーブル3']) === 'テーブル2',
  T.名前を決める(['テーブル1', 'テーブル3']));

console.log('\n[③ 1行目を 見出しに するか（実測の 見立て）]');
ok('①字だけ → 見出し', T.見出しか(['月', '売上', '原価']) === true);
ok('②数が 混ざる → 見出しでは ない', T.見出しか(['月', 100]) === false);
ok('②数だけ → 見出しでは ない', T.見出しか([1, 2]) === false);
ok('空だけ → 見出しでは ない', T.見出しか(['', '']) === false);
ok('★数の形の 字（"100"）も 数として 見る★', T.見出しか(['月', '100']) === false);

console.log('\n[④ 列の名（実測②＝列1・列2）]');
ok('見出しが 在れば その字', JSON.stringify(T.列名を作る(['月', '売上'], 2, true)) === '["月","売上"]',
  JSON.stringify(T.列名を作る(['月', '売上'], 2, true)));
ok('★見出しが 無ければ 列1・列2★', JSON.stringify(T.列名を作る(null, 2, false)) === '["列1","列2"]',
  JSON.stringify(T.列名を作る(null, 2, false)));
ok('空の 見出しは 列N で 埋める',
  JSON.stringify(T.列名を作る(['月', ''], 2, true)) === '["月","列2"]',
  JSON.stringify(T.列名を作る(['月', ''], 2, true)));
ok('★同じ名前は 使えない（後ろに 番号）★',
  JSON.stringify(T.列名を作る(['月', '月'], 2, true)) === '["月","月2"]',
  JSON.stringify(T.列名を作る(['月', '月'], 2, true)));

console.log('\n[⑤ しま（1つおきの 帯）]');
ok('中身の 1行目は 帯なし', T.しまか(0) === false);
ok('2行目は 帯', T.しまか(1) === true);
ok('3行目は 帯なし', T.しまか(2) === false);

console.log('\n[⑥ ★集計行＝左端 集計・右端だけ 式・間は 空（実測）★]');
{
  const 出 = T.集計行の中身(3, 'C2:C4');
  ok('3列 … 左端は 集計', 出[0] === '集計', JSON.stringify(出));
  ok('3列 … ★間は 空★', 出[1] === '', JSON.stringify(出));
  ok('3列 … 右端だけ SUBTOTAL(109,…)', 出[2] === '=SUBTOTAL(109,C2:C4)', JSON.stringify(出));
  const 二 = T.集計行の中身(2, 'B2:B4');
  ok('2列 … 集計 と 式だけ', JSON.stringify(二) === '["集計","=SUBTOTAL(109,B2:B4)"]', JSON.stringify(二));
  const 一 = T.集計行の中身(1, 'A2:A4');
  ok('1列 … 左端が 勝つ（集計）', JSON.stringify(一) === '["集計"]', JSON.stringify(一));
  ok('★9 では なく 109（隠した行を 数えない）★', T.集計の式('B2:B4').indexOf('109') > 0, T.集計の式('B2:B4'));
}

console.log('\n[⑦ どこか／重なり]');
{
  const t = { r1: 0, c1: 0, r2: 4, c2: 2, 見出し行: true, 集計行: true };
  ok('見出し', T.どこか(t, 0, 0) === '見出し');
  ok('中身', T.どこか(t, 1, 1) === '中身');
  ok('集計', T.どこか(t, 4, 2) === '集計');
  ok('外は null', T.どこか(t, 9, 9) === null);
  ok('★重なりを 見つける★', !!T.重なる([t], 3, 1, 6, 1));
  ok('離れていれば null', T.重なる([t], 10, 10, 12, 12) === null);
}

console.log('\n[⑧ 画面に つながっている]');
for (const n of ['表の箱', '今いる表', '表の窓を開く', '表のまわりを当てる', '表の範囲を読む',
  '表を決める', '集計行を出す', '集計行をしまう', '表をやめる', '表の塗り', '表の見出しか']) {
  ok(n + ' が 在る', !!抜く(n));
}
ok('窓が 在る', /id="tableOverlay"/.test(book));
ok('部品を 読み込んでいる', /src="lib\/table\.js/.test(book));
ok('★Ctrl+T で 出る（実Excelと 同じ）★', /if\(ek==='t'\)\{[^}]*表の窓を開く\(\)/.test(book));
ok('塗りの 所から 呼んでいる', /cell\.bgColor\|\|表の塗り\(r,c\)/.test(book));
ok('★手の 塗りが 勝つ（表の帯で 上書きしない）★',
  book.indexOf('cell.bgColor||表の塗り') > 0, '順番');
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑨ まわりを 当てる（空欄を 人に 埋めさせない）]');
{
  const f = new Function('sheets', 'activeSheet', 'ROWS', 'COLS',
    抜く('表のまわりを当てる') + '\nreturn 表のまわりを当てる;');
  const データ = {};
  for (let r = 0; r < 3; r++) for (let c = 0; c < 2; c++) データ[r + ',' + c] = { v: r + c + 1 };
  const 当てる = f([{ data: データ }], 0, 1000, 100);
  ok('中身の 塊を 取る', JSON.stringify(当てる(1, 1)) === '{"r1":0,"c1":0,"r2":2,"c2":1}',
    JSON.stringify(当てる(1, 1)));
  ok('★空の マスなら その1つだけ★', JSON.stringify(当てる(9, 9)) === '{"r1":9,"c1":9,"r2":9,"c2":9}',
    JSON.stringify(当てる(9, 9)));
}

console.log('\n[⑩ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { 表の窓を開く: function () { 受け = 'ok'; } };
  ACT['テーブル']();
  g.window = 前w;
  ok('「テーブル」→ 表の窓を開く', 受け === 'ok', String(受け));
}

console.log('\n[⑪ ★行・列を 足しても 何も 置いていかれない★]');
{
  /* ここが 抜けると 付箋が 別の行に 付き、リンクが 別のセルを 指す（★黙って ずれる★） */
  const 並び = book.slice(book.indexOf('var セルごとの棚 = ['), book.indexOf('function _moveValidations'));
  for (const 名 of ['validations', 'links', 'sparklines', 'comments']) {
    ok('棚の 一覧に 在る … ' + 名, 並び.indexOf("'" + 名 + "'") >= 0, 並び);
  }
  ok('テーブルも 動かしている', /_表を動かす\(s, 縦, 入れる, 上, 下\)/.test(book));
  const f = new Function('selR1', 'selR2', 'selC1', 'selC2',
    抜く('_表を動かす') + '\nreturn _表を動かす;');
  const 動かす = f(0, 0, 0, 0);
  {
    const s = { tables: [{ 名: 'A', r1: 5, c1: 0, r2: 8, c2: 2 }] };
    動かす(s, true, true, 2, 2);
    ok('上に 1行 入れたら 表も 1つ 下へ', s.tables[0].r1 === 6 && s.tables[0].r2 === 9,
      JSON.stringify(s.tables[0]));
  }
  {
    const s = { tables: [{ 名: 'A', r1: 5, c1: 0, r2: 8, c2: 2 }] };
    動かす(s, true, false, 2, 2);
    ok('上の 1行を 消したら 表も 1つ 上へ', s.tables[0].r1 === 4 && s.tables[0].r2 === 7,
      JSON.stringify(s.tables[0]));
  }
  {
    const s = { tables: [{ 名: 'A', r1: 5, c1: 0, r2: 8, c2: 2 }] };
    動かす(s, true, false, 4, 9);
    ok('★丸ごと 消したら 表も 消える★', s.tables.length === 0, JSON.stringify(s.tables));
  }
  {
    const s = { tables: [{ 名: 'A', r1: 5, c1: 0, r2: 8, c2: 2 }] };
    動かす(s, true, false, 6, 7);
    ok('★中を 2行 消したら 表は 2行 縮む★', s.tables[0].r1 === 5 && s.tables[0].r2 === 6,
      JSON.stringify(s.tables[0]));
  }
}

console.log('\n[⑫ ★見出しの ▼（自動フィルタ）★ 実測 ShowAutoFilter=True]');
{
  for (const n of ['見出しの印の場所', '見出しの印を描く', '見出しの印を押したか', '見出しの絞りを開く',
    '見出しの絞りを決める', '見出しの絞りを戻す', '見出しの絞りを全部', '見出しから並べ替え']) {
    ok(n + ' が 在る', !!抜く(n));
  }
  ok('窓が 在る', /id="headFilterOverlay"/.test(book));
  ok('★押した所から 呼んでいる（onMD）★', /見出しの印を押したか\(r, c, p\.x, p\.y\)/.test(book));
  ok('★描く所から 呼んでいる★', /見出しの印を描く\(sR, eR, sC, eC\)/.test(book));
  /* 場所は 描く所と 押す所で ★同じ物★ を 使う（ずれると 押せない ▼ に なる） */
  const f = new Function(抜く('見出しの印の場所') + '\nreturn 見出しの印の場所;')();
  const b = f(100, 200, 80, 22);
  ok('▼ は セルの 右はしに 出る', b.x + b.w <= 100 + 80, JSON.stringify(b));
  ok('▼ は セルの 高さに 収まる', b.y >= 200 && b.y + b.h <= 200 + 22, JSON.stringify(b));
  ok('★狭いセルでも 9px より 小さく しない★', f(0, 0, 20, 10).w >= 9, JSON.stringify(f(0, 0, 20, 10)));
  /* ★字と ▼ が 重ならない★（実ブラウザで 「売上」に 重なっていた） */
  ok('★見出しの 字は ▼ の 分だけ 狭める★',
    /var _印の幅 = 表の見出しか\(r,c\) \? \(見出しの印の場所\(x,y,w,h\)\.w \+ 6\) : 0;/.test(book));
  /* ★★切り取る 四角の 元が 変わりました★★（2026-09-24）
       前 ... `ctx.rect(x+1, y, max(1, w-2-_印の幅), h)`（★いつも マスの 四角★）
       今 ... `ctx.rect(_箱.x+1, y, max(1, _箱.w-2-_印の幅), h)`
              ＝★字が 入らない 時だけ 隣の 空いた マスまで 伸びる★（実Excel と 同じ）
       ★★ここで 守りたいのは 「▼ の 分を 引いて いるか」★★ なので、
       ★四角の 元（x か _箱.x か）は 見ません★＝★字の 形では なく 引き算を 見ます★ */
  ok('  その 幅で 切り取っている', /ctx\.rect\([^;]*-2-_印の幅\),h\)/.test(book));
  /* ★★`beginPath()` を 先に 呼んで いるか★★（2026-09-24 に 見つけた 穴）
       `rect()` は ★今の 道に 足す★／`clip()` は ★道 全部の 足し合わせ★ で 切る。
       ⇒呼ばないと ★前に 描いた 物の 道が 残り 切り取る 形が その時々で 変わる★。
       ⇒実測 ... 画面 1枚ぶん 描いた 後の 20行目が ★切れて いた★（`tests/mikire-webkit.mjs`） */
  ok('★★切り取る 前に 道を 空に して いる★★（`beginPath()`）',
    /ctx\.beginPath\(\);[^;]*[\s\S]{0,40}ctx\.rect\([^;]*-2-_印の幅\),h\)/.test(book));

  /* ★絞りは 部品の 正本を 使う（同じ物を 2つ 持たない）★ */
  const GF = require_(path.join(ROOT, 'lib/grid-filter.js'));
  ok('GridFilter.byValues が 在る', typeof GF.byValues === 'function');
  {
    const データ = { '0,0': { v: '月' }, '1,0': { v: '1月' }, '2,0': { v: '2月' }, '3,0': { v: '1月' } };
    const get = (r, c) => データ[r + ',' + c];
    const res = GF.byValues(get, { r1: 0, c1: 0, r2: 3, c2: 0 }, true, 0, { '1月': true });
    ok('★選んだ 値の 行だけ 残す★', JSON.stringify(res.keep) === '[1,3]', JSON.stringify(res.keep));
    ok('★見出し行は 隠さない★', res.hide.indexOf(0) < 0, JSON.stringify(res.hide));
    ok('残りは 隠す', JSON.stringify(res.hide) === '[2]', JSON.stringify(res.hide));
  }
  ok('★並べ替えは 正本（sortRange）を 呼ぶ★', /見出しから並べ替え[\s\S]{0,400}sortRange\(dir\)/.test(book));
  ok('★1つも ✓ が 無い時は 絞らない★', /1つも ✓ が ありません/.test(book));
  ok('★「作っていない物」の 一覧から ▼ を 外した★',
    !/★作っていない物★[^*]*見出しの▼ボタン/.test(book));
}

console.log('\ntable: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 集計を 9 に したら（実測は 109）気づけるか */
  if (T.集計の式('B2:B4').indexOf('SUBTOTAL(9,') >= 0) {
    素通り++; console.log('  ★素通り★ 9 に なっている（隠した行を 数えてしまう）');
  } else console.log('  ok   109 を 使っている');
  /* 壊し② 見出しの 見立てが「字が 1つでも 在れば 見出し」に なっていないか */
  if (T.見出しか(['月', 100]) === true) {
    素通り++; console.log('  ★素通り★ 数が 混ざっても 見出しに している');
  } else console.log('  ok   数が 混ざれば 見出しに しない');
  /* 壊し③ 棚の 一覧が 減っていないか */
  const 並び = book.slice(book.indexOf('var セルごとの棚 = ['), book.indexOf('function _moveValidations'));
  const 数 = (並び.match(/'/g) || []).length / 2;
  if (数 < 4) { 素通り++; console.log('  ★素通り★ 動かす棚が ' + 数 + '個に 減っている'); }
  else console.log('  ok   動かす棚は ' + 数 + '個');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
