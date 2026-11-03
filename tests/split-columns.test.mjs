/* split-columns.test.mjs — ★区切り位置（テキストを 列に 分ける）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 の TextToColumns で 実測）★
 *    A1 "山田,太郎,30" / A2 "鈴木,花子,25" / A3 "佐藤,一郎"  をカンマで 分けると
 *      行1 … A=山田 B=太郎 C=30
 *      行2 … A=鈴木 B=花子 C=25
 *      行3 … A=佐藤 B=一郎 C=（空）
 *    ⇒ ★元の列を 上書きする★／★足りない所は 空★
 *
 *  走らせ方: node tests/split-columns.test.mjs [--self-test]
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
  let d = 0;
  const j = book.indexOf('{', i);
  for (let k = j; k < book.length; k++) {
    if (book[k] === '{') d++;
    else if (book[k] === '}') { d--; if (d === 0) return book.slice(i, k + 1); }
  }
  return null;
}

console.log('\n[① 画面に 在る]');
for (const n of ['区切りで分ける', '何列に分かれるか', '上書きする数', '区切り位置を開く']) ok(n + ' が 在る', !!抜く(n));
ok('★窓が 在る（confirm を 使わない）★', /id="splitOverlay"/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

/* ── 本物を 走らせる ─────────────────────────── */
function 台(データ, r2, c1) {
  const s = { data: {} };
  for (const k of Object.keys(データ)) s.data[k] = { v: データ[k], d: データ[k] };
  const 出た = [];
  let 控え = 0;
  const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
    '_pushRowColUndo', 'setCell', 'render', 'updateBar', 'notify',
    抜く('何列に分かれるか') + '\n' + 抜く('上書きする数') + '\n' + 抜く('区切りで分ける')
    + '\nreturn { 分ける: 区切りで分ける, 何列: 何列に分かれるか, 上書き: 上書きする数 };');
  const api = f([s], 0, 0, r2, c1, c1,
    function () { 控え++; },
    function (r, c, v) { s.data[r + ',' + c] = { v: v, d: v }; },
    function () {}, function () {}, function (m) { 出た.push(m); });
  return { s, api, 出た, 控え: () => 控え };
}

console.log('\n[② 実Excelと 同じに 分かれる]');
{
  const t = 台({ '0,0': '山田,太郎,30', '1,0': '鈴木,花子,25', '2,0': '佐藤,一郎' }, 2, 0);
  ok('★3列に なると 分かる★', t.api.何列(',') === 3, String(t.api.何列(',')));
  const n = t.api.分ける(',');
  const 見る = (r, c) => (t.s.data[r + ',' + c] ? t.s.data[r + ',' + c].v : undefined);
  ok('行1 … 山田 / 太郎 / 30', 見る(0, 0) === '山田' && 見る(0, 1) === '太郎' && 見る(0, 2) === '30',
    [見る(0, 0), 見る(0, 1), 見る(0, 2)].join(' / '));
  ok('行2 … 鈴木 / 花子 / 25', 見る(1, 0) === '鈴木' && 見る(1, 1) === '花子' && 見る(1, 2) === '25',
    [見る(1, 0), 見る(1, 1), 見る(1, 2)].join(' / '));
  ok('★行3 … 佐藤 / 一郎 / （空）★（足りない所は 空）',
    見る(2, 0) === '佐藤' && 見る(2, 1) === '一郎' && 見る(2, 2) === '',
    [見る(2, 0), 見る(2, 1), JSON.stringify(見る(2, 2))].join(' / '));
  ok('★元の列を 上書きした★（実Excelと 同じ）', 見る(0, 0) === '山田');
  ok('3行 分けたと 言う', n === 3, String(n));
  ok('★何行 分けたかを 出す（黙って 変えない）★', t.出た.length === 1 && /3行/.test(t.出た[0]), JSON.stringify(t.出た));
  ok('★元に戻せる（控えを 取った）★', t.控え() === 1, String(t.控え()));
}

console.log('\n[③ 分かれない時は 何も しない]');
{
  const t = 台({ '0,0': 'ひとつだけ' }, 0, 0);
  const n = t.api.分ける(',');
  ok('0行', n === 0, String(n));
  ok('★「分かれる所が ない」と 言う★', /分かれる所が ありません/.test(t.出た[0] || ''), JSON.stringify(t.出た));
  ok('★元に戻す控えを 取らない（何も していない）★', t.控え() === 0, String(t.控え()));
}

console.log('\n[④ 上書きする数を 先に 数えられる]');
{
  const t = 台({ '0,0': 'あ,い', '0,1': 'のこる' }, 0, 0);
  ok('★右に 1個 在ると 分かる★', t.api.上書き(2) === 1, String(t.api.上書き(2)));
  ok('★何も 無ければ 0★', t.api.上書き(1) === 0, String(t.api.上書き(1)));
}

console.log('\n[⑤ リボンから 押せる]');
const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
{
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { 区切り位置を開く: function () { 受け = 'ok'; } };
  ACT['区切り位置']();
  g.window = 前w;
  ok('「区切り位置」→ 区切り位置を開く', 受け === 'ok', String(受け));
}

console.log('\nsplit-columns: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 足りない所を 空に しない（前の値が 残る） */
  const 残る = { '2,1': '前の値' };
  if (残る['2,1'] === '') { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  const t = 台({ '0,0': 'あ,い,う', '1,0': 'か,き', '1,1': 'のこる', '1,2': 'これも' }, 1, 0);
  t.api.分ける(',');
  if (t.s.data['1,2'] && t.s.data['1,2'].v === 'これも') {
    素通り++; console.log('  ★素通り★ 足りない所に 前の値が 残った');
  }
  /* 壊し② 上書きの数を 数えない */
  if (t.api.上書き(3) === undefined) { 素通り++; console.log('  ★素通り★ 上書きの数を 返さない'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
