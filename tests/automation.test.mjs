/* automation.test.mjs — ★自動化タブの 5つ（Office スクリプトの 見本と 同じ 中身）★ 2026-08-30
 *
 *  実Excelの「自動化」に 並ぶ 見本は ★Microsoftの クラウドで 動く TypeScript★。
 *  ★仕組みは 借りない★が ★やっている 事は うちでも 出来る★ので 中身を うちの 道具で 作った。
 *    ①選択範囲から サブテーブル ②ハイパーリンクを 全部 消す ③空の 行数を 数える
 *    ④テーブルを JSON で 出す ⑤テーブルから ピボット
 *  ボタンは 実Excelと 同じ 名前・同じ 並び（ギャラリー側と 一覧側の 2か所＝10個）。
 *
 *  走らせ方: node tests/automation.test.mjs [--self-test]
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
const bind = fs.readFileSync(path.join(ROOT, 'scripts/ribbon-bind.mjs'), 'utf8');

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

console.log('\n[① 5つとも 在る]');
for (const n of ['サブテーブルを作る', 'リンクを全部消す', '空の行を数える',
  'テーブルをJSONにする', '表からピボット']) ok(n + ' が 在る', !!抜く(n));
ok('★Microsoftの 仕組みは 借りないと 書いてある★',
  /仕組みは 借りないが ★やっている 事は うちでも 出来る★/.test(book));

console.log('\n[② ギャラリー側と 一覧側の 両方に 結んである（10個）]');
for (const 名 of ['選択範囲からサブテーブルを作成する', 'シートからハイパーリンクを削除する',
  '空の行数をカウントする', 'テーブル データを JSON として返す',
  'テーブルから新しいピボットテーブルを作成する']) {
  ok('ギャラリー側 … ' + 名, bind.indexOf("'自動化|Office スクリプト ギャラリー|" + 名 + "'") >= 0);
  ok('一覧側 … ' + 名, bind.indexOf("'自動化|Office スクリプト|" + 名 + "'") >= 0);
}

console.log('\n[③ 空の 行数を 数える（中身の 在る 一番下まで）]');
{
  const f = new Function('sheets', 'activeSheet', 'hideCtxMenu', 'notify',
    抜く('空の行を数える') + '\nreturn 空の行を数える;');
  const 言った = [];
  const 走 = (data) => f([{ data }], 0, () => {}, (m) => 言った.push(m))();
  ok('★間の 空だけ 数える（下の 果ては 数えない）★',
    走({ '0,0': { v: 'あ' }, '3,0': { v: 'い' } }) === 2, String(走({ '0,0': { v: 'あ' }, '3,0': { v: 'い' } })));
  ok('★空が 無ければ 0★', 走({ '0,0': { v: 'あ' }, '1,0': { v: 'い' } }) === 0);
  ok('★空文字は 中身と 見なさない★',
    走({ '0,0': { v: 'あ' }, '1,0': { v: '' }, '2,0': { v: 'う' } }) === 1,
    String(走({ '0,0': { v: 'あ' }, '1,0': { v: '' }, '2,0': { v: 'う' } })));
  const 前 = 言った.length;
  ok('★中身が 無ければ 0で 理由を 言う★', 走({}) === 0 && 言った.length > 前);
  ok('  数を 言葉でも 出す', /空の 行は/.test(言った.join('')), 言った.join(' / '));
}

console.log('\n[④ リンクを 全部 消す（★字は 消さない★）]');
{
  const f = new Function('sheets', 'activeSheet', 'hideCtxMenu', 'notify', 'render', 'updateBar',
    抜く('リンクを全部消す') + '\nreturn リンクを全部消す;');
  const s0 = {
    data: { '0,0': { v: 'リンクの字', color: '#467886', underline: true } },
    links: { '0,0': { 先: 'https://example.com/' } },
  };
  const 言った = [];
  const n = f([s0], 0, () => {}, (m) => 言った.push(m), () => {}, () => {})();
  ok('1個 消した', n === 1, String(n));
  ok('★字は 残る★', s0.data['0,0'].v === 'リンクの字', String(s0.data['0,0'].v));
  ok('★見た目（色・下線）は 元に 戻す★',
    s0.data['0,0'].color === undefined && s0.data['0,0'].underline === undefined,
    JSON.stringify(s0.data['0,0']));
  ok('★箱も 空に なる★', Object.keys(s0.links).length === 0);
  ok('★字は 消していない と 言う★', /字は 消していません/.test(言った.join('')), 言った.join(''));
  const s1 = { data: {}, links: {} };
  ok('リンクが 無ければ 0', f([s1], 0, () => {}, () => {}, () => {}, () => {})() === 0);
}

console.log('\n[⑤ サブテーブル／JSON／ピボット]');
{
  const 中 = 抜く('サブテーブルを作る') || '';
  ok('★サブテーブルは 新しい シートに 出す★',
    /var 前 = activeSheet;\s*addSheet\(\);/.test(中) && /setCell\(i2, j2, String\(v\), true\)/.test(中),
    中.slice(-200));
}
ok('★出したら テーブルに する★', /表の箱\(\)\.push\(\{ 名: 名, r1: 0, c1: 0/.test(book));
ok('★見出しの 見立ては 部品を 使う★', /Table\.見出しか\(一行目\)/.test(book));
ok('★中身が 空なら 断る★', /選んだ所に 中身が ありません/.test(book));
ok('★JSON は 1行目を 名前に する★', /名たち\.push\(見出し行 \?/.test(book));
ok('★数は 数の まま 出す★', /Pivot\._数にする\(v2\)/.test(book));
ok('★出す 中身が 無ければ 断る★', /出す 中身が ありません/.test(book));
ok('★ファイルの 出し方は 正本（FileOut）を 使う★',
  /FileOut\.deliver\(字, 名, \{ type: 'application\/json' \}\)/.test(book));
ok('★ここで Blob を 作らない（渡し口は 1本）★',
  !/new Blob\(\[字\]/.test(book));
ok('★渡し口が 無い時は 断る★', /ファイルの 渡し口が 読めていません/.test(book));
ok('★表から ピボットは 表の 中身だけ 選ぶ★', /if \(t\) sel\(t\.r1, t\.c1, t\.集計行 \? t\.r2 - 1 : t\.r2, t\.c2\)/.test(book));
ok('★alert / prompt / confirm を 使っていない★', !/\balert\(|\bprompt\(|\bconfirm\(/.test(book));

console.log('\n[⑥ リボンから 押せる]');
{
  const ACT = require_(path.join(ROOT, 'lib/ribbon-actions.js'));
  const g = globalThis, 前w = g.window;
  for (const [ボタン, 呼ぶ名] of [
    ['サブテーブルを作成', 'サブテーブルを作る'], ['ハイパーリンクを削除', 'リンクを全部消す'],
    ['空の行数をカウント', '空の行を数える'], ['テーブルをJSON', 'テーブルをJSONにする'],
    ['テーブルからピボット', '表からピボット'],
  ]) {
    let 受け = null;
    const w = {}; w[呼ぶ名] = function () { 受け = 'ok'; };
    g.window = w;
    ACT[ボタン]();
    g.window = 前w;
    ok('「' + ボタン + '」→ ' + 呼ぶ名, 受け === 'ok', String(受け));
  }
}

console.log('\nautomation: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝この 見張りが 見ている物を 直に 確かめる（★壊して 赤を 見るのは tools/break-check.mjs★）★');
  let 素通り = 0;
  const f = new Function('sheets', 'activeSheet', 'hideCtxMenu', 'notify', 'render', 'updateBar',
    抜く('リンクを全部消す') + '\nreturn リンクを全部消す;');
  const s0 = { data: { '0,0': { v: '字', color: '#467886', underline: true } }, links: { '0,0': {} } };
  f([s0], 0, () => {}, () => {}, () => {}, () => {})();
  if (s0.data['0,0'].v !== '字') { 素通り++; console.log('  ★素通り★ 字まで 消した'); }
  else console.log('  ok   字は 残る');
  if (s0.data['0,0'].underline) { 素通り++; console.log('  ★素通り★ 下線が 残っている'); }
  else console.log('  ok   下線も 外れる');
  const g2 = new Function('sheets', 'activeSheet', 'hideCtxMenu', 'notify',
    抜く('空の行を数える') + '\nreturn 空の行を数える;');
  const 数 = g2([{ data: { '0,0': { v: 'あ' }, '3,0': { v: 'い' } } }], 0, () => {}, () => {})();
  if (数 !== 2) { 素通り++; console.log('  ★素通り★ 空の行の 数が 合わない … ' + 数); }
  else console.log('  ok   空の行は 2');
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
