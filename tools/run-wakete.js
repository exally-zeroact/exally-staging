/* run-wakete.js — ★総なめを かたまりに 分けて 走らせ、最後に 1枚に まとめる★（2026-09-14）
 *
 *  ★★なぜ 要るか★★
 *    2026-09-14、`node tests/run.js` が ★9回中 6回★ system に
 *    「low on memory」で 殺されました（1回 11,300行・10分超）。
 *    実測（指示役1・経営者1）… メモリ 15.41GB／空き 3.4〜4.9GB／
 *    ★claude 12個 + node 24〜30個 ＝ 全体の 43〜47%★。
 *    ★本当の 食い主は 常駐の npx（1本 97MB × 5本以上）★＝席が 11 在る 分 積み上がる。
 *    ★特定の 試験が 悪いのでは ありません★＝2026-09-14 実測、
 *    2回とも 止まった `chuki.test.mjs --self-test` は ★単体なら 38秒で 完走★。
 *
 *  ★★決まりに 触れて いない事★★
 *    [[feedback_souname_wo_tochu_de_tomeruna]] の 原文が 言うのは ★2つだけ★：
 *      ①★途中で 止めるな★ ②★総なめは 1本だけ 走らせる★
 *    ★「分けるな」とは 1文字も 書いて いません★。
 *    止めたかったのは ★「緑だった所まで」で 報告する事★＝★見ていない分を 見た事に する★。
 *    ⇒ この 台は ★走らせた 本数を 記録に 残し、296本と 突き合わせて から★ 緑と 言います。
 *      ★1本でも 足りなければ 赤★。
 *
 *  ★★同じ 一覧・同じ 走らせ方★★
 *    `run.js` が 外に 出して いる `FILES` を ★そのまま★ 使います。
 *    ⇒★作る道を 2本に しない★（名簿を 2本に すると ★片方だけ 増えた試験★が 永久に 走らない）
 *
 *  使い方:
 *    node tools/run-wakete.js --hajime        … 記録を 消して 始める
 *    node tools/run-wakete.js 1 12            … 12に 分けた うちの 1つ目
 *    node tools/run-wakete.js --shime         … ★1枚に まとめて 出す（296本と 突き合わせ）★
 */
'use strict';
const { execFileSync } = require('child_process');
const fs = require('fs');
const path = require('path');
const { FILES } = require('../tests/run.js');

const 記録 = path.join(__dirname, '..', '.wakete-kiroku.json');
const 引数 = process.argv.slice(2);

if (!Array.isArray(FILES) || FILES.length < 10) {
  console.error('★run.js から 一覧が 読めない★（' + (FILES && FILES.length) + '本）');
  process.exit(2);
}
const 全部 = FILES.length;
const 名前 = (f) => { const [a, ...b] = Array.isArray(f) ? f : [f]; return a + (b.length ? ' ' + b.join(' ') : ''); };

const 読む = () => { try { return JSON.parse(fs.readFileSync(記録, 'utf8')); } catch (e) { return { 走った: {}, 落ちた: [] }; } };
const 書く = (o) => fs.writeFileSync(記録, JSON.stringify(o, null, 1) + '\n');

/* ══ ★始め★ ══ */
if (引数[0] === '--hajime') {
  /* ★★掛ける 前に 刻印を 見る★★（2026-09-14 に 足した）
     ★同じ日に 2度 踏みました★＝`lib/*.js` を 直したのに `?v=` を 貼り直さず、
     ★総なめの 1本目（stamp.test.mjs）で 落ちる★＝★12分割を 全部 やり直し★。
     ★「覚えて おく」では 3度目が 来ます★（指示役1 2026-09-14）ので ★道具に 入れます★。
     ⇒ 貼り直しが 要る なら ★ここで 止める★（走らせてから 気づくより 安い） */
  try {
    execFileSync(process.execPath,
      [path.join(__dirname, '..', 'scripts', 'stamp-build.mjs'), '--check'],
      { cwd: path.join(__dirname, '..'), stdio: 'pipe' });
  } catch (e) {
    console.error('');
    console.error('★★刻印が 古いです＝総なめを 始めません★★');
    console.error('  ⇒ `node scripts/stamp-build.mjs` を 走らせて から もう一度');
    console.error('  （JS/CSS を 直したら 毎回。★総なめの 1本目で どうせ 落ちます★）');
    process.exit(1);
  }
  書く({ 走った: {}, 落ちた: [], 始め: new Date().toISOString() });
  console.log('★刻印は 最新★（掛ける 前に 見ました）');
  console.log('★記録を 消して 始めました★（一覧 ' + 全部 + '本）');
  process.exit(0);
}

/* ══ ★★締め＝1枚に まとめる★★ ══ */
if (引数[0] === '--shime') {
  const k = 読む();
  const 走った本数 = Object.keys(k.走った).length;
  const 走って無い = FILES.map(名前).filter((n) => !k.走った[n]);
  console.log('');
  console.log('★★★総なめの 締め（分けて 走らせた 分を 1枚に）★★★');
  console.log('  一覧の 本数 ………… ★' + 全部 + '本★（run.js の FILES）');
  console.log('  走らせた 本数 ……… ★' + 走った本数 + '本★');
  console.log('  ★落ちた ……………… ' + k.落ちた.length + '本★');
  if (k.落ちた.length) k.落ちた.forEach((x) => console.log('     ・' + x));
  console.log('');
  if (走って無い.length) {
    console.log('★★回し忘れ ' + 走って無い.length + '本＝★緑に しません★★★');
    走って無い.slice(0, 20).forEach((x) => console.log('     ・' + x));
    if (走って無い.length > 20) console.log('     …ほか ' + (走って無い.length - 20) + '本');
    process.exit(2);
  }
  console.log('  ★全部 走らせた（回し忘れ 0本）★');
  console.log(k.落ちた.length ? '  ★★落ちが 在ります＝赤★★' : '  ★★落ち 0本＝緑★★');
  process.exit(k.落ちた.length ? 1 : 0);
}

/* ══ ★かたまりを 1つ 走らせる★ ══ */
const 何個目 = Number(引数[0]);
const 何分割 = Number(引数[1] || 12);
if (!何個目 || !何分割 || 何個目 < 1 || 何個目 > 何分割) {
  console.error('使い方: node tools/run-wakete.js <何個目> <何分割>  ／ --hajime ／ --shime');
  process.exit(2);
}
const 一かたまり = Math.ceil(全部 / 何分割);
const 組 = FILES.slice((何個目 - 1) * 一かたまり, 何個目 * 一かたまり);
if (!組.length) { console.log('★' + 何個目 + '/' + 何分割 + ' は 空です★'); process.exit(0); }

console.log('════ ' + 何個目 + ' / ' + 何分割 + ' … ' + 組.length + '本 ════');
const k = 読む();
let ng = 0;
for (const f of 組) {
  const [file, ...args] = Array.isArray(f) ? f : [f];
  const n = 名前(f);
  console.log('=== ' + n + ' ===');
  let 通った = true;
  try {
    execFileSync(process.execPath,
      ['--max-old-space-size=4096', path.join(__dirname, '..', 'tests', file), ...args],
      { stdio: 'inherit' });
  } catch (e) {
    通った = false; ng++;
    if (!k.落ちた.includes(n)) k.落ちた.push(n);
  }
  k.走った[n] = 通った ? 'ok' : 'ng';
  書く(k);                                  /* ★1本ごとに 書く★＝殺されても 記録が 残る */
}
console.log('');
console.log('★' + 何個目 + '/' + 何分割 + ' 終わり … 走らせた ' + 組.length + '本 ／ ★落ちた ' + ng + '本★');
console.log('  ここまでの 記録 … ' + Object.keys(k.走った).length + ' / ' + 全部 + '本');
console.log('  ★まだ 全部では ありません＝最後に `--shime` で 突き合わせて ください★');
process.exit(ng ? 1 : 0);
