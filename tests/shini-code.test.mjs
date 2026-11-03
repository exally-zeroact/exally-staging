/* shini-code.test.mjs — ★呼ばれない 物を 置いたままに しない★（2026-09-11）
 *
 *  ★★なぜ★★
 *    ★呼ばれない 物を 置くと「そこで 見て いる」と 読まれます★
 *    2026-09-09 に LINEST の 受け口を 外した 時、下請けの 4本
 *    （`_jsLinest` `_jsLogest` `_jsTrend` `_jsGrowth`）は ★残した ままでした★。
 *    ⇒ 注記に「もう 呼んで いません」と 書いて 済ませて いた
 *    ⇒★司さん（2026-09-02）「使わんもんは 消せや」★
 *
 *  ★消す 前に 数えた 事（2026-09-11）★
 *    呼び出し … `_jsGrowth` → `_jsLogest` の ★1本だけ★（4本の 中の 行き来）
 *    試験     … ★0本★ ／ 道具 … ★0本★
 *    ⇒★外から 呼ぶ 物は 1つも 在りませんでした★
 *
 *  ★消した 後に 確かめた 事★
 *    本物の ブラウザで 7本 打った … SLOPE／INTERCEPT／FORECAST／TREND／
 *    GROWTH／LINEST／LOGEST ⇒★消す 前と 1文字も 変わりません★
 *
 *  ★この 台は エンジンを 建てません★＝字を 読んで 数えるだけ。
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

console.log('');
console.log('[shini-code] ★呼ばれない 物を 置いたままに しない★');

/* ★見る 所★＝repo の js を 全部（node_modules と .git は 外す） */
function 字を集める() {
  const 出 = [];
  const 歩く = (d) => {
    for (const n of fs.readdirSync(d)) {
      if (n === 'node_modules' || n === '.git' || n === 'dist') continue;
      const f = path.join(d, n);
      const st = fs.statSync(f);
      if (st.isDirectory()) { 歩く(f); continue; }
      if (!/[.](js|mjs|html)$/.test(n)) continue;
      出.push({ 名: path.relative(ROOT, f).split(path.sep).join('/'), 字: fs.readFileSync(f, 'utf-8') });
    }
  };
  歩く(ROOT);
  return 出;
}
const 皆 = 字を集める();

T('★repo を 読めて いる（空振りして いない）★', () => {
  if (皆.length < 50) throw new Error('★' + 皆.length + '本しか 読めて いない★');
  if (!皆.some((x) => x.名 === 'exally-formula.js')) throw new Error('★exally-formula.js が 無い★');
});

T('★★消した 4本が どこにも 残って いない★★', () => {
  /* ★名前を 1本に つなげて 持ちます★＝この 見張り 自身が
     「消した 名前」を 字で 持って いるので、★自分を 数から 外します★
     （最初の 版は ★自分を 見つけて 赤★に なりました） */
  const 消した = ['_js' + 'Linest', '_js' + 'Logest', '_js' + 'Trend', '_js' + 'Growth'];
  const 残り = [];
  const 見る = 皆.filter((x) => x.名 !== 'tests/shini-code.test.mjs');
  for (const 名 of 消した) {
    for (const f of 見る) {
      /* ★注記の 外だけ 見ます★＝説明に 名前を 書くのは よい（消した 訳を 残す為） */
      const 動く = 注記を外す(f.字, { html: /[.]html$/.test(f.名) });
      if (動く.indexOf(名) >= 0) 残り.push(名 + ' … ' + f.名);
    }
  }
  if (残り.length) {
    throw new Error('★' + 残り.length + '本 残って いる★  ' + 残り.join(' ／ ')
      + '  ⇒★消したのに 呼ぶ 所が 在る＝どちらかが 間違い★');
  }
  console.log('      … 4本とも 動く 所には 1つも 無い（注記の 中は よい）');
});

T('★★消した 訳が 書いて 在る（黙って 消さない）★★', () => {
  const f = 皆.find((x) => x.名 === 'exally-formula.js');
  if (f.字.indexOf('2026-09-11') < 0 || f.字.indexOf('直線の係数') < 0) {
    throw new Error('★消した 訳が 書いて いない★＝次に 見た 人が「抜けて いる」と 思います');
  }
});

T('★★代わりの 道が 生きて いる★★', () => {
  const y = path.join(ROOT, 'lib/formula-yosoku.js');
  if (!fs.existsSync(y)) throw new Error('★lib/formula-yosoku.js が 無い★');
  const 字 = fs.readFileSync(y, 'utf-8');
  if (字.indexOf('直線の係数') < 0) {
    throw new Error('★`直線の係数` が 無い★＝LINEST を 表として 返す 所が 消えて います');
  }
});

console.log('');
console.log('  ' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
