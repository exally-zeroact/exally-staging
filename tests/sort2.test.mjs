/* sort2.test.mjs — ★並べ替え（2つの鍵）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 の Sort.SortFields で 実測）★
 *    名前/点 = ういろう30・あんこ20・かすてら30・ようかん10 を
 *    ★点↑ → 名前↑★ で 並べると:
 *      行1 見出し ／ 行2 ようかん10 ／ 行3 あんこ20 ／ ★行4 ういろう30 ／ 行5 かすてら30★
 *    ⇒ ★同じ点は 2つ目の鍵（名前）で 決まる★／★見出しの行は 動かない★
 *
 *  走らせ方: node tests/sort2.test.mjs [--self-test]
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

const G = require_(path.join(ROOT, 'lib/grid-sort.js'));
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const c = (v) => ({ v: v, d: v });

console.log('\n[① ★実Excelの 並びと 1行も 違わない★]');
{
  const 名 = [c('ういろう'), c('あんこ'), c('かすてら'), c('ようかん')];
  const 点 = [c(30), c(20), c(30), c(10)];
  const 順 = G.order2([{ cells: 点, dir: 'asc' }, { cells: 名, dir: 'asc' }]);
  const 出 = 順.map((i) => 名[i].v + 点[i].v).join(' / ');
  ok('点↑ → 名前↑', 出 === 'ようかん10 / あんこ20 / ういろう30 / かすてら30', 出);
}
console.log('\n[② 2つ目の鍵が 効いている（1つだけだと 違う並びに なる）]');
{
  const 名 = [c('ういろう'), c('あんこ'), c('かすてら'), c('ようかん')];
  const 点 = [c(30), c(20), c(30), c(10)];
  const 一つ = G.order2([{ cells: 点, dir: 'asc' }]).map((i) => 名[i].v).join(',');
  const 二つ = G.order2([{ cells: 点, dir: 'asc' }, { cells: 名, dir: 'asc' }]).map((i) => 名[i].v).join(',');
  ok('1つの鍵は 元の順を 保つ（ういろう が 先）', 一つ === 'ようかん,あんこ,ういろう,かすてら', 一つ);
  ok('★2つ目の鍵で 並びが 決まる★', 二つ === 'ようかん,あんこ,ういろう,かすてら', 二つ);
  const 逆 = G.order2([{ cells: 点, dir: 'asc' }, { cells: 名, dir: 'desc' }]).map((i) => 名[i].v).join(',');
  ok('★2つ目を 降順に すると 入れ替わる★', 逆 === 'ようかん,あんこ,かすてら,ういろう', 逆);
}
console.log('\n[③ 空白は いつも 最後（1つの鍵と 同じ決まり）]');
{
  const 点 = [c(30), c(''), c(10)];
  const 名 = [c('あ'), c('い'), c('う')];
  const 順 = G.order2([{ cells: 点, dir: 'asc' }, { cells: 名, dir: 'asc' }]);
  ok('空白が 最後', 順[順.length - 1] === 1, JSON.stringify(順));
  const 順2 = G.order2([{ cells: 点, dir: 'desc' }, { cells: 名, dir: 'asc' }]);
  ok('★降順でも 空白は 最後★', 順2[順2.length - 1] === 1, JSON.stringify(順2));
}
console.log('\n[④ 端の形]');
ok('鍵が 無ければ 空', JSON.stringify(G.order2([])) === '[]');
ok('鍵が null でも 落ちない', JSON.stringify(G.order2(null)) === '[]');
{
  const 順 = G.order2([{ cells: [c(1)], dir: 'asc' }]);
  ok('1行でも 動く', JSON.stringify(順) === '[0]', JSON.stringify(順));
}

console.log('\n[⑤ 画面に 繋がっている]');
ok('窓が 在る', /id="sortOverlay"/.test(book));
ok('開く働きが 在る', /function 並べ替えの窓を開く\(/.test(book));
ok('決める働きが 在る', /function 並べ替えを決める\(/.test(book));
ok('★order2 を 使っている★', /GridSort\.order2\(/.test(book));
ok('★見出しを 動かさない（先頭を ずらす）★', /var 先頭 = rng\.r1 \+ \(見出し \? 1 : 0\)/.test(book));
ok('★元に戻せる★', /_pushRowColUndo\(\);[\s\S]{0,400}GridSort\.order2|GridSort\.order2[\s\S]{0,400}_pushRowColUndo\(\)/.test(book));
ok('★何行 並べ替えたかを 言う★', /行を 並べ替えました/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
{
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { 並べ替えの窓を開く: function () { 受け = 'ok'; } };
  ACT['並べ替え']();
  g.window = 前w;
  ok('「並べ替え」→ 並べ替えの窓を開く', 受け === 'ok', String(受け));
}

console.log('\nsort2: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  /* 壊し① 2つ目の鍵を 見ない（1つ目だけ） */
  const 名 = [c('ういろう'), c('あんこ'), c('かすてら'), c('ようかん')];
  const 点 = [c(30), c(20), c(30), c(10)];
  const 逆 = G.order2([{ cells: 点, dir: 'asc' }, { cells: 名, dir: 'desc' }]).map((i) => 名[i].v).join(',');
  if (逆 === G.order2([{ cells: 点, dir: 'asc' }]).map((i) => 名[i].v).join(',')) {
    素通り++; console.log('  ★素通り★ 2つ目の鍵が 効いていない');
  }
  /* 壊し② 空白を 先に 出す */
  const 点2 = [c(30), c(''), c(10)];
  const 順 = G.order2([{ cells: 点2, dir: 'asc' }]);
  if (順[0] === 1) { 素通り++; console.log('  ★素通り★ 空白が 先に 来た'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
