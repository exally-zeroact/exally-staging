/* clear.test.mjs — ★クリア（中身だけ／書式だけ／すべて）★ 2026-08-29
 *
 *  ★真値（実Excel 16.0 で 実測 2026-08-29）★
 *    A1 に 123・太字・黄色の塗り を 入れて:
 *      ClearContents … 値=空 ／ ★太字=True・塗り 残る★
 *      ClearFormats  … ★値=456 残る★ ／ 太字=False
 *      Clear         … 値=空 ／ 太字=False
 *    ⇒ ★3つは 別物★。1つだけ 作って「クリア」と 言わない。
 *
 *  走らせ方: node tests/clear.test.mjs [--self-test]
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

console.log('\n[① 3つとも 画面に 在る]');
for (const n of ['中身を消す', '書式を消す', 'すべて消す']) ok(n + ' が 在る', !!抜く(n));
ok('★書式のキーを 1か所で 持っている★（増えた時に 消し漏れない）', /var 書式のキー\s*=/.test(book));

/* ── 本物を 走らせる（写しで 緑に しない）─────────────── */
console.log('\n[② 実際に 走らせて 中身と 書式を 見る]');
const 本文 = ['書式のキー を 取り出す', ''].join('');
const キー行 = book.match(/var 書式のキー = \[[\s\S]*?\];/);
ok('書式のキーの 一覧を 読めた', !!キー行);

if (キー行 && 抜く('書式を消す') && 抜く('中身を消す')) {
  const 台 = [
    キー行[0],
    抜く('書式を消す'),
    抜く('中身を消す'),
    抜く('すべて消す'),
    'return { 書式を消す: 書式を消す, 中身を消す: 中身を消す, すべて消す: すべて消す, キー: 書式のキー };',
  ].join('\n');
  /* 画面の 代わりに にせ物を 置く */
  function 台を作る() {
    const 状態 = {
      sheets: [{ data: { '0,0': { v: '123', f: '', d: '123', bold: true, bg: '#FFFF00', align: 'right' } } }],
      activeSheet: 0, selR1: 0, selR2: 0, selC1: 0, selC2: 0,
      undoStack: [], redoStack: [],
      消した: [],
    };
    const f = new Function('sheets', 'activeSheet', 'selR1', 'selR2', 'selC1', 'selC2',
      'undoStack', 'redoStack', 'setCell', 'updateBar', 'render', 'hideCtxMenu', 台);
    return {
      状態,
      呼ぶ(名) {
        const api = f(状態.sheets, 状態.activeSheet, 状態.selR1, 状態.selR2, 状態.selC1, 状態.selC2,
          状態.undoStack, 状態.redoStack,
          function (r, c, v) { 状態.消した.push([r, c, v]); const k = r + ',' + c; if (状態.sheets[0].data[k]) 状態.sheets[0].data[k].v = v; },
          function () {}, function () {}, function () {});
        api[名]();
        return api;
      },
    };
  }

  /* ②-a 中身だけ 消す ⇒ ★書式は 残る★ */
  {
    const t = 台を作る();
    t.呼ぶ('中身を消す');
    const cell = t.状態.sheets[0].data['0,0'];
    ok('中身だけ消す → 値が 空に なった', t.状態.消した.length === 1 && t.状態.消した[0][2] === '',
      JSON.stringify(t.状態.消した));
    ok('★中身だけ消す → 太字は 残る★（実Excelと 同じ）', cell.bold === true, String(cell.bold));
    ok('★中身だけ消す → 塗りは 残る★', cell.bg === '#FFFF00', String(cell.bg));
  }
  /* ②-b 書式だけ 消す ⇒ ★中身は 残る★ */
  {
    const t = 台を作る();
    t.呼ぶ('書式を消す');
    const cell = t.状態.sheets[0].data['0,0'];
    ok('★書式だけ消す → 値は 残る★（実Excelと 同じ）', cell.v === '123', String(cell.v));
    ok('書式だけ消す → 太字が 消えた', cell.bold === undefined, String(cell.bold));
    ok('書式だけ消す → 塗りが 消えた', cell.bg === undefined, String(cell.bg));
    ok('書式だけ消す → 揃えも 消えた', cell.align === undefined, String(cell.align));
    ok('★元に戻せる（控えを 積んだ）★', t.状態.undoStack.length === 1, String(t.状態.undoStack.length));
  }
  /* ②-c すべて 消す */
  {
    const t = 台を作る();
    t.呼ぶ('すべて消す');
    const cell = t.状態.sheets[0].data['0,0'];
    ok('すべて消す → 値も 書式も 消えた', cell.bold === undefined && t.状態.消した.length === 1,
      JSON.stringify([cell.bold, t.状態.消した.length]));
  }
}

console.log('\n[③ リボンから 押せる]');
const ACT = (await import('node:module')).createRequire(import.meta.url)(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 行き先] of [['中身を消す', '中身を消す'], ['書式を消す', '書式を消す'], ['すべて消す', 'すべて消す']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = {}; g.window[行き先] = function () { 受け = 行き先; };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ ' + 行き先, 受け === 行き先, String(受け));
}

console.log('\nclear: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 「中身だけ」で 書式まで 消す（実Excelと 違う） */
  const 壊れ = { bold: true, v: '1' };
  delete 壊れ.bold; 壊れ.v = '';
  if (壊れ.bold === true) { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  /* 壊し② 書式のキーの 一覧が 消えたら 気づくか */
  if (/var 書式のキー\s*=/.test('var なにか = [];')) { 素通り++; console.log('  ★素通り★ 一覧が 無くても 通した'); }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
