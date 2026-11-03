/* run-tests-batch.mjs — ★試験を 分けて 走らせる（一覧は tests/run.js 本人から 読む）★
 *
 *  ★なぜ在るか（2026-08-29）★
 *    通しで 走らせると 10分を 超えて 途中で 止められる事が ある。
 *    そこで 分けて 走らせるが、★一覧を 自分で 拾い直したら 7本 落として
 *    「全部 緑」と 嘘の報告を した★（CIが 赤で 捕まえた）。
 *    ⇒ ★一覧は 写さない。tests/run.js 本人から 読む★（require で 読める形に した）。
 *
 *  使い方:
 *    node scripts/run-tests-batch.mjs            ... 何本 在るかだけ 出す
 *    node scripts/run-tests-batch.mjs 1 61       ... 1本目〜61本目を 走らせる
 *    node scripts/run-tests-batch.mjs 1 61 --回  ... ★走り切った 時だけ★ 束の 記録を 書く（下の 断り）
 *
 *  ★★`--回`（2026-10-04・経営者と 決めた）★★
 *    ★穴★ ... 束が 途中で 切られる（timeout・メモリで 止められる）と 最後の 行が 出ない まま 終わり、
 *      読む 側が ★「赤 0」と 見分けられなかった★（10-04 に 1回 見間違えかけた）
 *    ⇒★読む 側を 道具に する★＝`scripts/souname-matome.mjs`
 *      ①まとめが 始めに 作る 回の 札（名札・頭・名簿の 指紋）を 読む
 *      ②★自分の 束の 古い 記録を 先に 消す★（前の 回の 完走を 「在る」と 数えない）
 *      ③走らせる 前と 後に ★木が 空か★ を 取る（自己試験が 途中で 落ちると 壊した 値が 残る）
 *      ④★走り切った 時だけ★ 書く＝★別名に 書いて 最後に 名前を 変える★（半端な 記録を 残さない）
 *      ⇒★途中で 切られた 束は 記録が 無い＝まとめが 赤★
 */
import path from 'node:path';
import fs from 'node:fs';
import crypto from 'node:crypto';
import { spawnSync } from 'node:child_process';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';
import { 回の置き場, 名簿の指紋, 頭を取る, 木を取る } from './souname-matome.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const { FILES } = require_(path.join(ROOT, 'tests/run.js'));

const 一覧 = FILES.map((f) => (Array.isArray(f) ? f : [f, null]));
const 数の引数 = process.argv.slice(2).filter((a) => a.slice(0, 2) !== '--');
const 回で = process.argv.includes('--回');
const 始 = Number(数の引数[0] || 0);
const 終 = Number(数の引数[1] || 0);
/* ★名簿の 本数を いつも 1行★（経営者＝CI と 比べる 時の 分母） */
if (始 && 終) console.log('★名簿 ' + 一覧.length + '本★（tests/run.js 本人から 読んだ）');
let 回 = null, 束の道 = null, 始めた = null, 前の木 = null;
if (回で && 始 && 終) {
  const 回の道 = path.join(回の置き場(), '回.json');
  if (!fs.existsSync(回の道)) { console.log('★回の 札が 無い★＝先に node scripts/souname-matome.mjs --hajime'); process.exit(2); }
  回 = JSON.parse(fs.readFileSync(回の道, 'utf8'));
  束の道 = path.join(回の置き場(), '束-' + 始 + '-' + 終 + '.json');
  if (fs.existsSync(束の道)) fs.unlinkSync(束の道);   /* ★古い 記録を 先に 消す★ */
  始めた = new Date().toISOString();
}
/* ★木は いつも 前と 後に 取る★（10-04・経営者の 気づき＝束の 出しだけ 見る 人には 汚れが 見えなかった） */
if (始 && 終) 前の木 = 木を取る(ROOT);
if (!始 || !終) {
  console.log('★試験は 全部で ' + 一覧.length + '本★（tests/run.js 本人から 読んだ）');
  console.log('  例: node scripts/run-tests-batch.mjs 1 61');
  process.exit(0);
}

/* ══ ★★「立ち上がれなかった」を 緑に 混ぜない★★ ══（2026-09-26・経営者1 の 知らせ）
     ★★何が 起きたか（経営者1 の 手元）★★
       機械の 空きが 足りず ★318本が `3221225794`（0xC0000142）を 返した★
       ＝Windows の ★「プロセスを 作れなかった」★
       ＝★試験が 落ちた のでは なく 1度も 走らなかった★
       ⇒★★「通った」とも 「割れた」とも 言えません★★
     ★★この 道具は 前から 捕まえます★★＝`r.status !== 0` は 全部 赤
       ⇒★3221225794 も 赤に なります★（字で 確かめました）
     ★★それでも 足りない 物が 1つ 在りました★★
       ＝★「赤 0件」だけでは 何本 走らせたか 出て いません★
       ⇒★だから 走らせた 本数と 緑の 本数を いつも 出します★
       ⇒★立ち上がれなかった 時は その 名を 書きます★
     ★記憶★ [[feedback_ci_node_version_vs_engines]]（落ちるのでは なく 走らない） */
var 立ち上がれない = { 3221225794: '★プロセスを 作れなかった（0xC0000142・機械の 空き 不足）★' };
let 赤 = 0, 走った = 0, 緑 = 0;
const 落ちた = [];
for (let i = 始 - 1; i < Math.min(終, 一覧.length); i++) {
  const [f, a] = 一覧[i];
  const args = ['--max-old-space-size=4096', path.join(ROOT, 'tests', f)];
  if (a) args.push(a);
  const r = spawnSync(process.execPath, args, { encoding: 'utf8' });
  走った++;
  if (r.status !== 0) {
    赤++;
    var 訳 = r.signal ? '（★中で殺された signal=' + r.signal + '★）'
      : (立ち上がれない[r.status] ? '（' + 立ち上がれない[r.status] + '）'
        : (r.error ? '（★走らせられません ' + r.error.message + '★）'
          : '（自分で ' + r.status + ' を返した）'));
    落ちた.push((i + 1) + '本目 ' + f + (a ? ' ' + a : '') + 訳);
  } else { 緑++; }
}
/* ★分母を 必ず 出します★＝★「赤 0件」だけでは 何本 走らせたか 分かりません★ */
const 頼んだ = Math.min(終, 一覧.length) - (始 - 1);
console.log((始) + '〜' + Math.min(終, 一覧.length) + '本目 … ★赤 ' + 赤 + '件★'
  + '（★頼んだ ' + 頼んだ + '本 ／ 走らせた ' + 走った + '本 ／ 緑 ' + 緑 + '本★）');
/* ★★木の 汚れを 束の 出しにも 出す★★（10-04・経営者の 気づき）
     前は ★束の 出しは 木を 判じず★「赤 0件」と 言い、汚れを 赤と 言うのは まとめ だけ だった
     ⇒★`--回` の 時は 汚れなら 束も 赤（exit 1）★／`--回` 無しは 出す だけ（手元で 作りかけを 試す 時の 為） */
const 後の木 = 木を取る(ROOT);
const 木の字 = (s) => (s ? '★汚れ★' : '空');
console.log('   木：前 ' + 木の字(前の木) + ' ／ 後 ' + 木の字(後の木)
  + ((前の木 || 後の木) ? (回 ? '  ⇒★木が 汚れて いる＝この 束は 赤★' : '  ⇒（--回 無し＝判じない）') : ''));
/* ★走り切った ここで だけ 記録を 書く★（別名に 書いて 最後に 名前を 変える） */
if (回) {
  const 記録 = {
    名札: 回.名札, 頭: 頭を取る(ROOT), 名簿の指紋: 名簿の指紋(FILES), 名簿: 一覧.length,
    始: 始, 終: Math.min(終, 一覧.length), 頼んだ: 頼んだ, 走らせた: 走った, 緑: 緑, 赤: 落ちた,
    始めた: 始めた, 終えた: new Date().toISOString(), 前の木: 前の木, 後の木: 後の木,
  };
  const 仮 = 束の道 + '.' + crypto.randomBytes(4).toString('hex') + '.kaki';
  fs.writeFileSync(仮, JSON.stringify(記録, null, 1));
  fs.renameSync(仮, 束の道);
}
if (走った !== 頼んだ) {
  console.log('   ★★頼んだ 本数と 走らせた 本数が 違います＝この 出しは 当てに なりません★★');
  process.exit(1);
}
for (const s of 落ちた) console.log('   ・' + s);
process.exit((赤 || (回 && (前の木 || 後の木))) ? 1 : 0);
