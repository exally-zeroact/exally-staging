/* valign.test.mjs — ★セルの 上下揃え（上/中央/下）★ 2026-08-29
 *
 *  ★なぜ在るか★
 *    リボンに「上揃え・上下中央揃え・下揃え」を 出す時、
 *    画面が ★valign を 1か所も 見ていなかった★（book.html に 0件）。
 *    そのまま 出していたら ★押しても 何も 変わらないボタン★に なっていた。
 *    ⇒ 画面（canvas）に 上下の位置を 入れた。ここは その見張り。
 *
 *  ★真値（実Excel 16.0 日本語版で 実測 2026-08-29）★
 *    ・★既定は 中央★（私は 最初「下揃え」と 思い込んでいた）
 *        COM … A1／シート全体／行／標準スタイル とも VerticalAlignment = -4108（xlCenter）
 *        保存した xlsx の 中身 … <alignment vertical="center"/>
 *    ・上 = -4160（xlTop） ／ 下 = -4107（xlBottom）
 *
 *  走らせ方: node tests/valign.test.mjs [--self-test]
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

/* ── 本物の _vY を book.html から 取り出して 走らせる（写しで 緑に しない）── */
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
const 本文 = 抜く('_vY');
ok('★book.html に _vY が 在る（試験が 空振りしていない）', !!本文);
const _vY = 本文 ? new Function(本文 + '\nreturn _vY;')() : null;

/* セル … y=100 / 高さ h=40 / 字の大きさ fs=12（行の高さ lh = 16.8） */
const y = 100, h = 40, 字 = 12, lh = 字 * 1.4;

console.log('\n[① 上・中央・下で 位置が 変わる]');
if (_vY) {
  const 上 = _vY('top', y, h, 字, 1);
  const 中 = _vY('middle', y, h, 字, 1);
  const 下 = _vY('bottom', y, h, 字, 1);
  ok('★3つとも 違う位置★（押して 何も 変わらない、が 起きない）',
    上 !== 中 && 中 !== 下 && 上 !== 下, [上, 中, 下].join(' / '));
  ok('上 < 中央 < 下 の順', 上 < 中 && 中 < 下, [上, 中, 下].join(' < '));
  ok('★上は セルの 上のふち＋余白★', Math.abs(上 - (y + 3 + lh / 2)) < 0.01, String(上));
  ok('★中央は 真ん中★', Math.abs(中 - (y + (h - lh) / 2 + lh / 2)) < 0.01, String(中));
  ok('★下は セルの 下のふち－余白★', Math.abs(下 - (y + h - 3 - lh + lh / 2)) < 0.01, String(下));

  console.log('\n[② はみ出さない]');
  ok('上は セルの 中', 上 - lh / 2 >= y, String(上));
  ok('下は セルの 中', 下 + lh / 2 <= y + h, String(下));
  console.log('\n[③ 折り返し（何行でも）]');
  const 上3 = _vY('top', y, h, 字, 3);
  const 下3 = _vY('bottom', y, h, 字, 3);
  ok('★行が 増えても 上は 動かない★', Math.abs(上3 - 上) < 0.01, String(上3));
  ok('★行が 増えると 下は 上へ 寄る★', 下3 < 下, 下3 + ' < ' + 下);
  ok('★背が 足りない時も はみ出さない★', _vY('bottom', y, 10, 字, 5) >= y, String(_vY('bottom', y, 10, 字, 5)));
}

console.log('\n[④ 画面が 実際に 使っている]');
ok('★既定は 中央★（実Excel 日本語版の 実測どおり）', /valign\s*=\s*cell\.valign\s*\|\|\s*'middle'/.test(book));
ok('★1行の 描画で 使っている★', /ctx\.fillText\(display,\s*x\+w\/2,\s*vy\)/.test(book));
ok('★折り返しの 描画でも 使っている★', /startY2\s*=\s*_vY\(/.test(book));
/* ★セルの字を 描く所★だけ 見る。行番号の見出し（fillText(r+1,…)）は セルの字では ない。
   ★除く物は 名指しで 書く★＝黙って 数から 落とさない。 */
const 決め打ち = (book.match(/fillText\([^)]*,\s*y\s*\+\s*h\s*\/\s*2\s*\)/g) || []);
const セルの字 = 決め打ち.filter((x) => !/fillText\(r\s*\+\s*1/.test(x));
ok('★セルの字で y+h/2 の 決め打ちが 残っていない★', セルの字.length === 0, セルの字.join(' / '));
ok('★除いたのは 行番号の見出し 1件だけ★', 決め打ち.length - セルの字.length === 1, String(決め打ち.length));

console.log('\n[⑤ リボンから 押せる]');
const ACT = (await import('node:module')).createRequire(import.meta.url)(path.join(ROOT, 'lib/ribbon-actions.js'));
for (const [名, 期待] of [['上揃え', 'top'], ['上下中央', 'middle'], ['下揃え', 'bottom']]) {
  const g = globalThis, 前w = g.window;
  let 受け = null;
  g.window = { applyFormat: function (k, v) { 受け = [k, v]; } };
  ACT[名]();
  g.window = 前w;
  ok('「' + 名 + '」→ valign=' + 期待, JSON.stringify(受け) === JSON.stringify(['valign', 期待]), JSON.stringify(受け));
}

console.log('\nvalign: ' + 緑 + '/' + (緑 + 赤) + ' passed');

if (壊す) {
  console.log('\n★--self-test＝わざと 壊して 赤に なるか★');
  let 素通り = 0;
  /* 壊し① 上下を 見ない（前の姿＝y+h/2 の 決め打ち） */
  const 壊れ = function (v, y2, h2) { return y2 + h2 / 2; };
  if (壊れ('top', y, h) !== 壊れ('bottom', y, h)) { 素通り++; console.log('  ★素通り★ 壊し方が おかしい'); }
  if (_vY && _vY('top', y, h, 字, 1) === _vY('bottom', y, h, 字, 1)) {
    素通り++; console.log('  ★素通り★ 上と 下が 同じ位置なのに 通した');
  }
  /* 壊し② 既定を 下に 戻す（思い込みの 姿） */
  if (/valign\s*=\s*cell\.valign\s*\|\|\s*'bottom'/.test(book)) {
    素通り++; console.log('  ★素通り★ 既定が 下に 戻っている（実測は 中央）');
  }
  if (素通り) { console.log('★抜け道 ' + 素通り + '件★'); process.exit(1); }
  console.log('  ok   抜け道 0件');
}
process.exit(赤 ? 1 : 0);
