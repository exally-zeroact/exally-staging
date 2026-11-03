/* make-prompt.mjs — ★AIの 頭を 台帳から 機械で 作る★（2026-09-05 司さん）
 *
 *  ★司さん★「★AIの構造は ファイルにして 更新できるような 設計に しとけよ★」
 *
 *  ★なぜ 要るか（実測で 分かった 事）★
 *    api/claude.js に ★手書きの 使えない関数リスト 22個★が 埋まっていた。
 *    正本（lib/formula-extra.js の 数える()）と 突き合わせたら ★17個 間違い★：
 *      ・★足してあるのに「使えない」と 教えていた … 4個★
 *        TOCOL / TOROW / CHOOSEROWS / CHOOSECOLS（★今日 私が 直した 物★）
 *      ・★動かないのに 1つも 教えていない … 13個★
 *        LAMBDA / LET / MAP / REDUCE / SCAN / BYROW / BYCOL / MAKEARRAY /
 *        STOCKHISTORY / FIELDVALUE / IMAGE / PHONETIC / YEN
 *        ★★YEN は この 数え方が 間違いだった（同じ日の 夜に 実測で 分かった）★★
 *          ＝★打てば ¥1,235 が 出る★（Exally が 本名 DOLLAR に 直している）
 *          ⇒★「動かない」ではなく ★別名で動く★の 棚へ 移した★
 *          ⇒ もし 直していなかったら ★AIが 客に「YEN は 動きません」と 言っていた★
 *      ・しかも 前置きは ★「LET・LAMBDA等を 積極的に 使用」と 名指しで 勧めていた★
 *    ⇒★人が 写すから ずれる★。★写さない形に する★。
 *
 *  ★やる事★
 *    lib/formula-extra.js の 台帳 → prompt/kansuu.md を 作り直す
 *    ★手で 書かない★（このファイルを 走らせて 出す）
 *
 *  使い方: node scripts/make-prompt.mjs          … 作り直す
 *          node scripts/make-prompt.mjs --check  … ★ずれていたら 赤★（CIが 見る）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const 出す先 = path.join(ROOT, 'prompt/kansuu.md');

const 台帳 = require_(path.join(ROOT, 'lib/formula-extra.js')).数える();
const 足す = 台帳.足す.slice().sort();
const 別lib = Object.keys(台帳.別のlibで足す || {}).sort();
const 別名 = Object.keys(台帳.別名で動く || {}).sort();
const 形で = Object.keys(台帳.形で動く || {}).sort();
const 足さない = 台帳.足さない;

/* ★日付を AIへ 渡す 所から 外す★＝日付が 混ざると ★毎回 置き直し＝毎回 置き賃★。
   理由の 全文は 台帳(lib/formula-extra.js)に 残っている。 */
function 日付を外す(s) {
  return String(s).replace(new RegExp('（?20\\d\\d-\\d\\d-\\d\\d[^）]*）?', 'g'), '').trim();
}

function 作る() {
  const L = [];
  L.push('# ★Exally で 使える 関数・使えない 関数★');
  L.push('');
  L.push('> ★この ファイルは 機械が 作ります。手で 書かないで ください★');
  L.push('> 作り方 … `node scripts/make-prompt.mjs`');
  L.push('> 正本 …… `lib/formula-extra.js` の `数える()`');
  L.push('> ★手で 写していた 頃は 22個中 17個 間違っていました（2026-09-05 実測）★');
  L.push('');
  L.push('> ★★AIへ 渡るのは 下の ``` の 中だけ★★（日付を 中に 入れない＝置き賃が 毎回 かかる）');
  L.push('');
  L.push('```');   /* ★ここから 下が AIへ 渡る★ */
  L.push('## ★足した 関数（Exally で 動く）★');
  L.push('');
  L.push(足す.join(' / '));
  if (別lib.length) {
    L.push('');
    L.push('★別の 部品で 足した 物★');
    for (const f of 別lib) L.push('- ' + f + ' … ' + 台帳.別のlibで足す[f]);
  }
  /* ★★別名で 動く 物は ★動く 側★に 出す★★（2026-09-05 に 足した）
     ★なぜ★ … YEN が 「足さない」に 置いてあった
       ⇒ AIへ 渡る 紙で ★「Exally内では まだ 動かない」の 列に 並んでいた★
       ⇒★実際は ¥1,235 が 出るのに AIが お客さんに「動きません」と 言う★所だった
     ⇒★★「足していない」と「動かない」は 別★★ */
  if (別名.length) {
    L.push('');
    L.push('★別の 名前で 動く 物（★打たれたら そのまま 動く＝断っては いけない★）★');
    for (const f of 別名) L.push('- ' + f + ' … ' + 日付を外す(台帳.別名で動く[f]));
  }
  /* ★★3つ目の 棚＝『形に よって 出る』★★（2026-09-08）
     ★AI は これを「動く」とも「動かない」とも 言っては いけない★
     ⇒★★『この 形なら 出ます／この 形は まだです』と 言う★★
     ★なぜ 別の 列に するか★
       ・「動く」に 混ぜると ★お客さんの 式が 黙って 止まる★
       ・「動かない」に 置くと ★測って 分かった 事を 1つも 渡さない★ */
  if (形で.length) {
    L.push('');
    L.push('★形に よって 出る 物（★「動く」とも「動かない」とも 言わない★）★');
    L.push('- ★この 棚の 物は 下の 説明の とおり 答えて ください★');
    L.push('  ⇒「★この 形なら 出ます／この 形は まだです★」と 形を 名指しで 言う');
    for (const f of 形で) L.push('- ' + f + ' … ' + 日付を外す(台帳.形で動く[f]));
  }
  L.push('');
  L.push('## ★★Exally で 動かない 関数（★勧めては いけない★）★★');
  L.push('');
  L.push('★聞かれたら「Exally内では まだ 動かない」と はっきり 言う。★');
  L.push('★代わりの やり方を 出す。★★黙って 勧めない★★');
  L.push('');
  for (const f of Object.keys(足さない).sort()) {
    L.push('- ★' + f + '★ … ' + 日付を外す(足さない[f]));
  }
  L.push('```');
  L.push('');
  return L.join('\n');
}

const 新 = 作る();
if (process.argv.includes('--check')) {
  const 今 = fs.existsSync(出す先) ? fs.readFileSync(出す先, 'utf8') : '';
  if (今 !== 新) {
    console.log('✗ prompt/kansuu.md が 台帳と ずれています');
    console.log('  → `node scripts/make-prompt.mjs` を 走らせて commit してください');
    console.log('  ★人が 書き足すと AIが 嘘を 言います（2026-09-05 に 17個 ずれていた）★');
    process.exit(1);
  }
  console.log('✓ prompt/kansuu.md は 台帳と 一致（動く ' + (足す.length + 別lib.length + 別名.length)
    + '個 / 動かない ' + Object.keys(足さない).length + '個）');
  process.exit(0);
}
fs.mkdirSync(path.dirname(出す先), { recursive: true });
fs.writeFileSync(出す先, 新);
console.log('作り直した … prompt/kansuu.md（動く ' + (足す.length + 別lib.length + 別名.length)
  + '個 / 動かない ' + Object.keys(足さない).length + '個）');
