/* ribbon-context.test.mjs — ★コンテキストタブ（物を 選んだ時だけ 出る タブ）★ 2026-08-31
 *
 *  ★なぜ★（司さん 2026-08-30）
 *    「Excel を 細胞分解レベルまで 網羅して 把握した上で 持ち込み パクる」
 *    ★新規の 空ブックだけ 見ていると この 8タブ 235部品は 1つも 見えない★。
 *    2026-08-30 に 司さんの 問い「構造まで 深く リサーチもしたよな？」で 気づき、
 *    表・グラフ・図形・ピボットを 作って 選んで 取り直した。
 *
 *  ★実測（docs/excel-ribbon-context-2026-08-30.tsv）★
 *      図形の書式 55 ／ 書式 47 ／ ピボットテーブル分析 28 ／ テーブル デザイン 26
 *      スライサー 24 ／ タイムライン 23 ／ デザイン 16 ／ グラフのデザイン 16
 *      ＝★8タブ／235部品★
 *
 *  ★取る時に 踏んだ 罠（記録）★
 *    ★グラフの 2タブは 1回目 失敗した★＝ホームの 中身が 出た。
 *    ★タブを 選べたか（IsSelected）を 確かめてから 中身を 取る★形に 直した。
 *
 *  ★この 見張りが 見る物★
 *    ① 正本は 自動生成（手で 書いていない）
 *    ② 8タブ／235部品が 減っていない
 *    ③ ★どの 物を 選んだら どのタブが 出るか★が 全タブに 付いている
 *    ④ 並びが 実測の tsv と ★1行ずつ 同じ★
 *    ⑤ つないだ数を 数えられる（★数えられない物を「出来ている」と 言わない★）
 *
 *  走らせ方: node tests/ribbon-context.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const NL = String.fromCharCode(10), TAB = String.fromCharCode(9);
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const 正本 = require_(path.join(ROOT, 'lib/ribbon-context-spec.js'));
const 働き = require_(path.join(ROOT, 'lib/ribbon-actions.js'));

console.log('★コンテキストタブ（物を 選んだ時だけ 出る タブ）★');

/* ── ① 手で 書いていない ── */
const 生道 = path.join(ROOT, 'docs/excel-ribbon-context-2026-08-30.tsv');
言う(fs.existsSync(生道), '★元の 実測が 在る★');
const 字 = fs.readFileSync(path.join(ROOT, 'lib/ribbon-context-spec.js'), 'utf8');
言う(/自動生成/.test(字) && /excel-ribbon-context-2026-08-30\.tsv/.test(字),
  '★正本は 自動生成（元が 書いてある）★', '★手で 書くと 実Excel と ずれる★');

/* ── ② 数 ── */
const 数 = 正本.数える(働き);
言う(数.タブ === 8, '★8タブ★（今 ' + 数.タブ + '）');
言う(数.部品 === 235, '★235部品★（今 ' + 数.部品 + '）');

const 期待 = { '図形の書式': 55, '書式': 47, 'ピボットテーブル分析': 28, 'テーブル デザイン': 26,
  'スライサー': 24, 'タイムライン': 23, 'デザイン': 16, 'グラフのデザイン': 16 };
for (const 名 of Object.keys(期待)) {
  const T = 正本.ツリー.find((v) => v.name === 名);
  const n = T ? T.groups.reduce((s, g) => s + g.items.length, 0) : 0;
  言う(n === 期待[名], '　' + 名 + ' … ' + 期待[名] + '部品', '今 ' + n);
}

/* ── ③ 出る 条件 ── */
const 条件なし = 正本.ツリー.filter((v) => !v.出る).map((v) => v.name);
言う(条件なし.length === 0, '★どの タブにも「何を 選んだら 出るか」が 付いている★',
  条件なし.join(' / '));
言う(正本.出るタブ('グラフ').length === 2,
  '★グラフを 選ぶと 2タブ（グラフのデザイン／書式）★',
  正本.出るタブ('グラフ').map((v) => v.name).join(' / '));
言う(正本.出るタブ('ピボット').length === 2,
  '★ピボットを 選ぶと 2タブ（ピボットテーブル分析／デザイン）★');
言う(正本.出るタブ('テーブル').length === 1 && 正本.出るタブ('図形').length === 1,
  '★表・図形は それぞれ 1タブ★');
言う(正本.出るタブ('在るわけない物').length === 0, '★知らない 物には タブを 出さない★');

/* ── ④ 並びが 実測と 1行ずつ 同じ ── */
const 生 = fs.readFileSync(生道, 'utf8').split(NL)
  .filter((l) => l.trim() && !l.startsWith('#')).map((l) => l.split(TAB));
const 平ら = 正本.全部();
言う(平ら.length === 生.length,
  '★行の 数が 実測と 同じ★（正本 ' + 平ら.length + ' / 実測 ' + 生.length + '）');
let ずれ = 0, 最初 = '';
for (let i = 0; i < Math.min(平ら.length, 生.length); i++) {
  const a = 平ら[i], b = 生[i];
  if (a.t !== (b[0] || '').trim() || a.g !== (b[1] || '').trim() || a.p !== (b[2] || '').trim()) {
    ずれ++;
    if (!最初) 最初 = i + '行目: 正本[' + a.t + '|' + a.g + '|' + a.p + '] 実測[' + b.join('|') + ']';
  }
}
言う(ずれ === 0, '★並びが 実測と 1行ずつ 同じ★', 最初);

/* ── ⑤ つないだ数を 数えられる ── */
言う(typeof 数.つないだ === 'number' && typeof 数.まだ === 'number',
  '★つないだ数を 数えられる（今 ' + 数.つないだ + '／' + 数.部品 + '）★',
  '★数えられない物を「出来ている」と 言わない★');
言う(数.つないだ + 数.まだ === 数.部品, '★つないだ ＋ まだ ＝ 全部★');
言う(数.つないだ >= 25, '★名前が 一致する 働きが 25個 以上 在る（今 ' + 数.つないだ + '）★');

/* ── わざと 壊して 赤に なるか ── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  言う(正本.数える(null).つないだ === 0, '★働きを 渡さなければ 0（数え過ぎない）★');
  言う(正本.数える({}).つないだ === 0, '★働きが 空なら 0★');
  const T = 正本.ツリー[0];
  言う(!!T.name && Array.isArray(T.groups) && T.groups.length > 0,
    '★タブの 形が 壊れていない（名前と 組が 在る）★');
  const g = T.groups[0];
  言う(!!g.name && Array.isArray(g.items) && g.items.length > 0,
    '★組の 形が 壊れていない（名前と 部品が 在る）★');
  言う(正本.全部().every((v) => v.t && v.g && v.p),
    '★どの 部品にも タブ・組・名前が 在る★');
}

console.log('\nribbon-context: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
