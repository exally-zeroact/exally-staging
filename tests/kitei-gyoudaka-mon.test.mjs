/* ★★既定の 行の 高さの 門★★（2026-11-07）
   板の 既定の 行の 高さ（customHeight）を 知るのは ★1か所★ に 寄せた
     ・画面＝book.html の `既定の行高(sh)`（と `素の行高(sh, r)`）
     ・読む＝js/book-open.js の 変数 `既定の行の点` を 板と 図形の 置き場の 両方へ
   ★1か所ずつ 答えの 試験を 当てると 穴が 残った★（設計前の 叩きで 3回 止められた）
     ＝答えが 変わらない 所（yToR の 当たり）・取り消しの 元・引っ張りの 始まり・窓 は 答えの 試験で 殺せない
   ⇒★字で 数える★：画面の 既定の 名 `ROW_H` を 書いて よいのは 下の 3行だけ
   ⇒ 図形の 置き場に 行の 高さを 数字で 直に 渡したら 赤
   自己確認＝写しを わざと 壊して 門に 通し 赤を 見る（試験の 中で 毎回） */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const 根 = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const 本 = fs.readFileSync(path.join(根, 'book.html'), 'utf8');
const 開く = fs.readFileSync(path.join(根, 'js/book-open.js'), 'utf8');

/* ★許す 行（字の 完全一致・前後の 空白は 見ない）★ */
const 許す = [
  '*   scrollTop, scrollLeft, HDR_W, HDR_H, COL_W, ROW_H, COLS, ROWS',
  'var ROW_H = 24, COL_W = 72, HDR_W = 46, HDR_H = 22, ROWS = 1048576, COLS = 16384;',
  'function 既定の行高(sh){ return (sh && sh.既定の行高) || ROW_H; }',
];

function 門(本の字, 開くの字) {
  const 外れ = [];
  const 見た = new Set();
  本の字.split('\n').forEach((行, i) => {
    if (!/\bROW_H\b/.test(行)) return;
    const 字 = 行.trim();
    if (許す.includes(字)) { 見た.add(字); return; }
    外れ.push('book.html:' + (i + 1) + ' ' + 字.slice(0, 90));
  });
  for (const 字 of 許す) if (!見た.has(字)) 外れ.push('★許す 行が 無い★ ' + 字);
  /* 図形の 置き場＝台に載せる形(図形, colW, 既定の列幅, rowH, 既定の行高, …) の 第5引数 */
  const 呼び = [...開くの字.matchAll(/台に載せる形\(([^)]*)\)/g)];
  if (!呼び.length) 外れ.push('★js/book-open.js に 台に載せる形( が 無い★');
  for (const m of 呼び) {
    const 第5 = (m[1].split(',')[4] || '').trim();
    /* ★\b は 日本語の 後ろで 効かない★＝空白か 終わりで 切る */
    if (!/^既定の行の点(\s|$)/.test(第5)) 外れ.push('js/book-open.js 台に載せる形 の 行の 既定＝「' + 第5 + '」（既定の行の点 で 渡す）');
  }
  return 外れ;
}

let pass = 0, fail = 0;
const ok = (名, 真, 訳) => { if (真) { pass++; console.log('  ✓ ' + 名); } else { fail++; console.log('  ✗ ' + 名 + (訳 ? ' — ' + 訳 : '')); } };

console.log('\n[kitei-gyoudaka-mon] ★既定の 行の 高さを 知るのは 1か所★');
const 今 = 門(本, 開く);
ok('★ROW_H は 3行だけ・図形の 置き場は 既定の行の点★', 今.length === 0, 今.slice(0, 5).join(' ／ '));

/* ★自己確認＝わざと 壊して 赤（毎回 回す）★ */
const 壊し方 = [
  { 名: '7209 歩く数が 板の 既定を 見ない', 先: '本', 前: '  var _低い = 既定の行高(_sh), _細い', 後: '  var _低い = ROW_H, _細い' },
  { 名: '11834 取り消しの 元が 板の 既定を 見ない', 先: '本', 前: '  var old = 素の行高(sheets[activeSheet], r);', 後: '  var old = sheets[activeSheet].rowH[r] || ROW_H;' },
  { 名: '図形の 置き場に 24 を 直に 渡す', 先: '開く', 前: 'rowH, 既定の行の点 || 24, hiddenCols', 後: 'rowH, 24, hiddenCols' },
];
for (const こ of 壊し方) {
  const 元 = こ.先 === '本' ? 本 : 開く;
  if (元.split(こ.前).length !== 2) { ok(こ.名 + ' → 赤', false, '★壊す 所が 1つ で ない★（形が 変わった＝見張りを 直す）'); continue; }
  const 壊 = 元.replace(こ.前, () => こ.後);
  const 外 = こ.先 === '本' ? 門(壊, 開く) : 門(本, 壊);
  ok(こ.名 + ' → 赤', 外.length > 0, '★緑のまま（門が 空振り）★');
}

console.log('\n  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
