/* formula-bar-type.test.mjs — ★数を 入れても 数式バーが 落ちない★ 2026-08-30
 *
 *  ★見つけ方（実ブラウザ・Playwright 2026-08-30）★
 *    スパークラインを 実UIで 試そうと して setCell(r,c,10) と 数を 入れたら:
 *      TypeError: (cell.f || "").replace is not a function
 *        at updateBar → sel
 *    ＝f（打った字）に ★数が そのまま★ 入り、字の道具（replace）が 使えなかった。
 *    ★落ちるのは 数式バーだけでは ない★＝sel() の中で 投げるので
 *      ★セルを 選ぶ そのものが 出来なくなる★。
 *
 *  ★直し★ f は 必ず 字で 持つ（v は 打った物の まま＝計算は 変えない）。
 *
 *  走らせ方: node tests/formula-bar-type.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 壊す = process.argv.includes('--self-test');
let 緑 = 0, 赤 = 0;
const ok = (名, 条件, 添え) => {
  if (条件) { 緑++; console.log('  ok   ' + 名); }
  else { 赤++; console.log('  ★NG★ ' + 名 + (添え !== undefined ? '  … ' + 添え : '')); }
};
const book = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');

console.log('\n[① 入れる時に 字に する]');
{
  const i = book.indexOf('var next = Object.assign({}, existing, {v:v, f:');
  ok('セルを 作る所が 見つかる', i >= 0);
  const 行 = book.slice(i, book.indexOf('\n', i));
  ok('★f は String() を 通している★', /f:\(v===null\|\|v===undefined\)\?'':String\(v\)/.test(行.replace(/\s/g, '')), 行);
  ok('★v は そのまま（計算を 変えない）★', /\{v:v,/.test(行.replace(/\s/g, '')), 行);
}

console.log('\n[② 実際に 通して 落ちないか]');
{
  /* setCell の 中の「作る所」だけ 取り出して 数を 通す */
  const 作る = (v) => {
    const d = v;
    const existing = {};
    return Object.assign({}, existing, { v: v, f: (v === null || v === undefined) ? '' : String(v), d: String(d) });
  };
  for (const [名, v, 期待] of [
    ['数', 10, '10'], ['マイナスの数', -5, '-5'], ['0', 0, '0'],
    ['字', 'あ', 'あ'], ['式', '=A1+1', '=A1+1'],
    ['null', null, ''], ['undefined', undefined, ''],
    ['真', true, 'true'],
  ]) {
    const cell = 作る(v);
    let 落ちた = null;
    try { (cell.f || '').replace(/\n/g, ' '); } catch (e) { 落ちた = String(e && e.message); }
    ok('数式バーが 落ちない … ' + 名, 落ちた === null, String(落ちた));
    ok('  f = ' + JSON.stringify(期待) + ' … ' + 名, cell.f === 期待, JSON.stringify(cell.f));
  }
  /* ★v は 数の まま★（合計が 字に なると 黙って 小さくなる） */
  ok('★v は 数の まま★', typeof 作る(10).v === 'number', typeof 作る(10).v);
}

console.log('\n[③ 打った字を 読み返す所も 字で 比べている]');
{
  /* 数式バーで 直した時、stored と 比べて 同じなら 入れ直さない。
     stored が 数だと ★いつも 違う★に なって 毎回 入れ直していた。 */
  ok('比べている所が 在る', /if\(actual!==stored\)/.test(book));
  ok('★f が 字なので 比べられる★（①で String を 通している）', /String\(v\)/.test(book));
}

console.log('\nformula-bar-type: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝直す前の 書き方に 戻したら 赤に なるか★');
  let 素通り = 0;
  const 昔 = Object.assign({}, {}, { v: 10, f: 10, d: '10' });
  let 落ちた = false;
  try { (昔.f || '').replace(/\n/g, ' '); } catch (e) { 落ちた = true; console.log('  ok   直す前の 持ち方は ちゃんと 落ちる … ' + e.message); }
  if (!落ちた) { 素通り++; console.log('  ★素通り★ 数を f に 入れても 落ちなかった＝この試験は 何も 見ていない'); }
  if (!/String\(v\)/.test(book)) { 素通り++; console.log('  ★素通り★ String() が 消えている'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
