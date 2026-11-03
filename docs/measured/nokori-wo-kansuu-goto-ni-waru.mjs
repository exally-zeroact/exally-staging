/* ★「借り物だけが 答える 分」を 関数ごとに 割る★（2026-09-15）
 *
 *  ★なぜ 要るか★（指示役1 2026-09-15）
 *    ★「外して 悪く なる 行 ＝ 10通り」を 見出しに しては いけない★
 *    ＝★借り物だけが 答える 4,645通り★は ★外した 瞬間に #NAME? に なる★
 *    ⇒★本当の 代金は そちら★。★では その 4,645 は どの 関数が 作って いるのか★を 割る。
 *
 *  ★★借り物を 立てません★★
 *    `ryouhou-ni-keisan-saseru.mjs` は ★式ごとに 借り物を 立て直す★＝25分 かかります。
 *    ここは ★名前を 見るだけ★＝★うち（土台7＋皮61）が 知らない 名前★を 数えます。
 *    ⇒★同じ 数に なるとは 限りません★（向こうは ★実際に 押した 結果★／ここは ★名前だけ★）
 *      ・向こうが 多い … 借り物が 名前を 知って いても 答えを 出せない 形が 在る
 *      ・こちらが 多い … うちが 名前を 知って いても 答えを 出せない 形が 在る
 *    ⇒★数が ずれたら それ自体が 見所★です。★どちらも 隠さず 出します★。
 *
 *  使い方: node docs/measured/nokori-wo-kansuu-goto-ni-waru.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const R = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..') + path.sep;
const require_ = createRequire(path.join(R, 'package.json'));
const K = require_(path.join(R, 'lib/shiki-kansuu.js'));
const T = require_(path.join(R, 'lib/shiki-tsunagi.js'));
const Z = await import('file:///' + path.join(R, 'tests/zairyou.mjs').replace(/\\/g, '/'));

const 出せる = new Set([...Object.keys(K.表 || {}), ...T.名前たち()]);

/* ★実Excel の 名簿★（借り物に 無い 物も 在る） */
const 実Excel = new Set(fs.readFileSync(path.join(R, 'docs/measured/excel-functions-2026-09-06.txt'), 'utf8')
  .split('\n').map((s) => s.trim()).filter(Boolean));

const 判定の字 = /^(合った|違う|こちらだけ誤り|名前が通らない|数えない|同じ|一致)$|直った|変わらず（前から/;
const 紙 = [];
(function なめる(d) {
  for (const e of fs.readdirSync(d, { withFileTypes: true })) {
    const p = path.join(d, e.name);
    if (e.isDirectory()) { なめる(p); continue; }
    if (!/\.(tsv|csv)$/.test(e.name)) continue;
    if (/^cases-/.test(e.name) || /-awase-/.test(e.name)) continue;
    紙.push(p);
  }
})(path.join(R, 'docs'));

const 数 = {};            /* 関数 → 式の 本数（うちが 知らない 物だけ） */
let 全式 = 0, 知っている = 0;
for (const p of 紙) {
  let t;
  try { t = fs.readFileSync(p, 'utf8'); } catch (e) { continue; }
  const 生 = t.split('\n');
  if (生.some((l) => l && !l.startsWith('#') && l.split('\t').some((v) => 判定の字.test((v || '').trim())))) continue;
  const 見た = new Set();
  for (const l of 生) {
    if (!l || l.startsWith('#')) continue;
    for (const v of l.split('\t')) {
      const x = (v || '').trim();
      if (!/^=[A-Z]/.test(x)) continue;
      if (見た.has(x)) break;
      見た.add(x);
      全式++;
      const 外 = Z.外の関数(x);
      if (!外) break;
      if (出せる.has(外)) { 知っている++; break; }
      数[外] = (数[外] || 0) + 1;
      break;
    }
  }
}

const 並 = Object.keys(数).sort((a, b) => 数[b] - 数[a]);
const 合 = 並.reduce((s, k) => s + 数[k], 0);

console.log('# ★「うちが まだ 知らない 関数」が 何本の 式を 作って いるか★（2026-09-15）');
console.log('');
console.log('★数え方★（★借り物を 立てて いません＝名前を 見るだけ★）');
console.log('  ・紙 … `docs` の tsv/csv（記録の 紙・cases-・-awase- は 外す）… ' + 紙.length + '本');
console.log('  ・式 … ★同じ 式は 1回だけ★ … ' + 全式 + '通り');
console.log('  ・うちが 出せる … 土台 ' + Object.keys(K.表 || {}).length + '個 ＋ 皮 ' + T.名前たち().length + '個');
console.log('');
console.log('★★うちが まだ 知らない 関数が 作る 式 … ' + 合 + '通り★★（' + 並.length + '個の 関数）');
console.log('  うちが 知っている 関数の 式 … ' + 知っている + '通り');
console.log('');
console.log('★関数ごと（多い順・全部）★');
let 積 = 0;
for (const k of 並) {
  積 += 数[k];
  console.log('  ' + String(数[k]).padStart(5) + '本  ' + k.padEnd(18)
    + (実Excel.has(k) ? '実Excel に 在る' : '★実Excel に 無い★')
    + '   （ここまで ' + 積 + '本／' + Math.round(積 / 合 * 100) + '%）');
}
console.log('');
const 無 = 並.filter((k) => !実Excel.has(k));
console.log('★実Excel に 無い 名前★ … ' + 無.length + '個 = ' + (無.join(' ') || 'なし'));
console.log('  ＝★真似る 相手が 居ない★＝★書かなくて よい かもしれません★（要 判じ）');
