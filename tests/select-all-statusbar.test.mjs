/* select-all-statusbar.test.mjs — ★すべて選ぶ（Ctrl+A）で 帯が 落ちない★ 2026-08-30
 *
 *  ★見つけ方（実測）★
 *    リボンを増やした後の 実UI掃き出しで、キーを押す所から こう出た:
 *      RangeError: Invalid array length
 *        at Array.push … updateStatusBar → updateBar → sel → onKD
 *    ＝Ctrl+A は sel(0,0,ROWS-1,COLS-1)＝★1,048,576 × 16,384 ＝ 17,179,869,184 マス★。
 *      これを 1つずつ 配列に 積んでいた。配列の長さの上限（2^32-1）を 超えて 死ぬ。
 *    ★死ぬのは 帯だけでは ない★＝sel() の中で 投げるので
 *      ★「すべて選ぶ」そのものが 何も起きない★（render() まで 届かない）。
 *
 *  ★直し方★
 *    GridStats は ★空マスを 数えない★（grid-stats.js 実測④）ので、
 *    ★中身が 入っているマスだけ★ 拾えば 出る答えは 1つも 変わらない。
 *
 *  走らせ方: node tests/select-all-statusbar.test.mjs [--self-test]
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
const GridStats = require_(path.join(ROOT, 'lib/grid-stats.js'));

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

/* ★実物の 大きさ★＝book.html から 読む（決め打ちしない） */
const m = book.match(/ROWS\s*=\s*(\d+)\s*,\s*COLS\s*=\s*(\d+)/);
const ROWS = m ? +m[1] : 0, COLS = m ? +m[2] : 0;
console.log('\n[① 実物の 大きさ（book.html から 読んだ）]');
ok('ROWS × COLS が 読めた', ROWS > 0 && COLS > 0, ROWS + ' × ' + COLS);
ok('★配列の上限（2^32-1）を 超える大きさ★＝素朴に積むと 必ず 死ぬ',
  ROWS * COLS > 4294967295, String(ROWS * COLS));
ok('Ctrl+A は すべてを 選ぶ', /if\(ek==='a'\)\{[^}]*sel\(0,0,ROWS-1,COLS-1\)/.test(book));

/* 画面の関数を そのまま 走らせる（読むだけでは 落ちるか 分からない） */
function 帯を作る(data, r1, c1, r2, c2) {
  const 出た = { html: null, 高さ: 0 };
  const el = { set innerHTML(v) { 出た.html = v; }, get innerHTML() { return 出た.html; } };
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selC1', 'selR2', 'selC2',
    'document', 'GridStats', '_sbHeight',
    抜く('updateStatusBar') + '\nreturn updateStatusBar;');
  f([{ data }], 0, r1, c1, r2, c2,
    { getElementById: (id) => (id === 'status-bar' ? el : null) },
    GridStats, () => { 出た.高さ++; })();
  return 出た.html;
}

console.log('\n[② ★すべて選んでも 落ちない★]');
{
  const data = { '0,0': { v: 10 }, '5,3': { v: 20 }, '900000,9000': { v: 30 }, '2,2': { v: 'あ' } };
  let 落ちた = null, h = null;
  try { h = 帯を作る(data, 0, 0, ROWS - 1, COLS - 1); }
  catch (e) { 落ちた = String(e && e.message); }
  ok('★Ctrl+A で 落ちない★', 落ちた === null, String(落ちた));
  ok('合計 60 が 出る（10+20+30）', /合計/.test(h || '') && />60</.test(h || ''), String(h));
  ok('データの個数 4 が 出る（文字も 数える）', />4</.test(h || ''), String(h));
}

console.log('\n[③ ★答えが 前と 1つも 変わらない★（小さい範囲で 突き合わせ）]');
{
  /* 素朴に 全マス 積む やり方＝★直す前の 答え★ を その場で 作り、同じかどうか 見る */
  const 作る = (data, r1, c1, r2, c2) => {
    const cells = [];
    for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) cells.push(data[r + ',' + c] || null);
    return JSON.stringify(GridStats.items(GridStats.summarize(cells)));
  };
  const 例 = [
    ['数だけ', { '0,0': { v: 10 }, '0,1': { v: 20 }, '1,0': { v: 30 } }, 0, 0, 3, 3],
    ['文字まじり', { '0,0': { v: 10 }, '0,1': { v: 'あ' }, '2,2': { v: 20 } }, 0, 0, 4, 4],
    ['数の形の文字', { '0,0': { v: '10', f: "'10" }, '0,1': { v: 5 } }, 0, 0, 2, 2],
    ['エラー入り', { '0,0': { v: 10 }, '0,1': { f: '=1/0', d: '#DIV/0!' } }, 0, 0, 2, 2],
    ['空だけ', {}, 0, 0, 5, 5],
    ['1マスだけ', { '3,3': { v: 7 } }, 3, 3, 3, 3],
    ['★範囲の外は 入れない★', { '0,0': { v: 10 }, '9,9': { v: 999 } }, 0, 0, 2, 2],
  ];
  for (const [名, data, r1, c1, r2, c2] of 例) {
    const 前 = 作る(data, r1, c1, r2, c2);
    const 今 = 帯を作る(data, r1, c1, r2, c2);
    const 今の中身 = JSON.stringify(GridStats.items(GridStats.summarize(
      Object.keys(data).filter((k) => {
        const p = k.indexOf(','), kr = +k.slice(0, p), kc = +k.slice(p + 1);
        return kr >= r1 && kr <= r2 && kc >= c1 && kc <= c2;
      }).map((k) => data[k]))));
    ok('同じ答え … ' + 名, 前 === 今の中身, 前 + ' ／ ' + 今の中身);
    ok('  帯も 出る … ' + 名, 今 !== null, String(今));
  }
}

console.log('\n[④ ★大きく選んでも 速い★]');
{
  const data = {}; for (let i = 0; i < 5000; i++) data[i + ',0'] = { v: i };
  const t = process.hrtime.bigint();
  帯を作る(data, 0, 0, ROWS - 1, COLS - 1);
  const ms = Number(process.hrtime.bigint() - t) / 1e6;
  ok('★Ctrl+A が 1秒 かからない★', ms < 1000, ms.toFixed(1) + 'ms');
}

console.log('\nselect-all-statusbar: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝直す前の 書き方に 戻したら 赤に なるか★');
  let 素通り = 0;
  const data = { '0,0': { v: 10 } };
  /* ★★直す前の 書き方を ★子どもの プロセス★ で 走らせる★★（2026-08-31）
   *
   *  ★なぜ 分けるか（実測）★
   *    直す前の 書き方は ★17,179,869,184 マスを 積む★＝わざと 死なせる 検査。
   *    手元（記憶が 多い）では `RangeError: Invalid array length` が 出て 終わる。
   *    ★CI では 例外の 前に 機械ごと 殺される★（実測＝SIGTRAP／2026-08-31）。
   *    ⇒ ★親を 巻き込む★ので ★別の プロセスで 走らせて 死んだ事だけ 見る★。
   *    ★死に方（例外か 強制終了か）は 機械で 変わる★ので ★どちらでも 合格★に する。
   *    ★中身は 前と 同じ 書き方★＝手加減して 小さくは していない。
   */
  const { spawnSync } = await import('node:child_process');
  const 素朴 = 'const data={"0,0":{v:10}};const cells=[];'
    + 'for(let r=0;r<=' + (ROWS - 1) + ';r++){for(let c=0;c<=' + (COLS - 1) + ';c++)'
    + 'cells.push(data[r+","+c]||null);}console.log("死ななかった");';
  const 子 = spawnSync(process.execPath, ['--max-old-space-size=256', '-e', 素朴],
    { encoding: 'utf8', timeout: 120000 });
  const 落ちた = 子.status !== 0 || !!子.signal;
  if (落ちた) {
    console.log('  ok   直す前の 書き方は ちゃんと 死ぬ … '
      + (子.signal ? ('殺された signal=' + 子.signal)
        : ((子.stderr || '').split('\n').find((l) => /Error/.test(l)) || ('exit=' + 子.status)).trim()));
  } else { 素通り++; console.log('  ★素通り★ 直す前の 書き方でも 死ななかった＝この試験は 何も 見ていない'); }
  /* 今の画面が 素朴な 書き方に 戻っていないか */
  const いま = 抜く('updateStatusBar') || '';
  if (!/Object\.keys\(data\)/.test(いま)) { 素通り++; console.log('  ★素通り★ 中身の入ったマスだけ 拾う書き方が 消えている'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
