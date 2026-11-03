/* ribbon-scope.test.mjs — ★空の箱を (a)(b)(c) に 分けている★ 2026-08-30
 *
 *  ★なぜ★（監査役が 絵3を 見て 見つけた・2026-08-30）
 *    空の箱 9組 ★全部★ が 同じ「これから」だった。
 *    その うち ★7組（16個）は「作らない」と 決めた 物★＝
 *    ★守れない 約束を お客さんに 見せていた★。
 *    うちの決まり「★出来ていない物の ボタンを 見せるな★」に 反する。
 *
 *  ★3通り（監査役の 決め）★
 *    (a) これから 作る          → 「これから」
 *    (b) Excelに 在るが 付けない → ★「付けません」＋理由1行★
 *    (c) そもそも 押す物では ない → ★組ごと 出さない★
 *
 *  ★見る物★
 *    ① (a)(b)(c) の 数を 出し、★合計が 空の箱の数と 合う★
 *    ② ★(b) には 必ず 理由が 付いている★（空文字を 許さない）
 *    ③ ★使っていない 理由の行が 残っていない★（死んだ許可＝次を 見逃す）
 *    ④ ★Sheet1 タブは 描かない★（＝下の シート見出しの 写り込み）
 *
 *  走らせ方: node tests/ribbon-scope.test.mjs  ／  --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let ok = 0, ng = 0;
const 言う = (よい, 文, 添え) => {
  if (よい) { ok++; console.log('  ok   ' + 文); }
  else { ng++; console.log('  NG   ' + 文); if (添え) console.log('       ' + 添え); }
};

const 範囲 = (await import('file://' + path.join(ROOT, 'lib/ribbon-scope.js').replace(/\\/g, '/'))).default;
const specSrc = fs.readFileSync(path.join(ROOT, 'lib/ribbon-spec.js'), 'utf8');

const 部品 = [];
for (const 行 of specSrc.split(String.fromCharCode(10))) {
  const m = /\{ t: '([^']*)', g: '([^']*)', p: '([^']*)', a: (.*)\},?\s*$/.exec(行.trim());
  if (m) 部品.push({ t: m[1], g: m[2], p: m[3], 結んだ: m[4].trim() !== 'null' });
}
console.log('★空の箱の 分け方★');
言う(部品.length >= 280, '★空振りしていない★（部品 ' + 部品.length + '個 を 読めた）');

/* ── 組ごとに 分ける ───────────────────────────── */
const 組 = new Map();
for (const v of 部品) {
  const k = v.t + '|' + v.g;
  if (!組.has(k)) 組.set(k, []);
  組.get(k).push(v);
}
const 空 = [...組.entries()].filter(([, vs]) => !vs.some((v) => v.結んだ));

const A = [], B = [], C = [], 理由なし = [];
for (const [k, vs] of 空) {
  const [t, g] = k.split('|');
  if (範囲.出さない理由(t, g)) { C.push(k); continue; }
  const 理由たち = vs.map((v) => 範囲.付けない理由(t, g, v.p));
  if (理由たち.every((r) => r)) {
    B.push(k);
    if (理由たち.some((r) => !String(r).trim())) 理由なし.push(k);
  } else {
    A.push(k);
  }
}
console.log('  ── 実測 ──');
console.log('   (a) これから   … ' + A.length + '組  ' + A.join(' / '));
console.log('   (b) 付けません … ' + B.length + '組  ' + B.join(' / '));
console.log('   (c) 出さない   … ' + C.length + '組  ' + C.join(' / '));

言う(A.length + B.length + C.length === 空.length,
  '★(a)+(b)+(c) が 空の箱の数と 合う＝' + 空.length + '組★',
  A.length + B.length + C.length + ' ≠ ' + 空.length);

言う(理由なし.length === 0, '★(b) には 全部 理由が 付いている★', 理由なし.join(' / '));

/* ★「これから」と 出る物に「作らないと 決めた物」が 混ざっていない★ */
const 混ざり = [];
for (const k of A) {
  const [t, g] = k.split('|');
  for (const v of 組.get(k)) if (範囲.付けない理由(t, g, v.p)) 混ざり.push(k + '|' + v.p);
}
言う(混ざり.length === 0,
  '★「これから」の 箱に 作らないと 決めた物が 混ざっていない★（守れない 約束を 見せない）',
  混ざり.join(' / '));

/* ── ③ 使っていない 理由の行が 残っていないか ──── */
const 使った外 = new Set();
for (const v of 部品) if (範囲.付けない理由(v.t, v.g, v.p)) 使った外.add(v.t + '|' + v.g + '|' + v.p);
const 死んだ外 = Object.keys(範囲.対象外).filter((k) => !使った外.has(k));
言う(死んだ外.length === 0,
  '★対象外の 行が 全部 使われている＝' + Object.keys(範囲.対象外).length + '行★',
  '★まだ 正本に 無い物★: ' + 死んだ外.join(' / '));

const 使ったC = new Set(C);
const 死んだC = Object.keys(範囲.出さない).filter((k) => !使ったC.has(k));
言う(死んだC.length === 0,
  '★出さないの 行が 全部 使われている＝' + Object.keys(範囲.出さない).length + '行★',
  死んだC.join(' / '));

/* ── ④ 実際に 描いて 見る（Sheet1 が 出ない・(b)に 字が 出る）── */
const dom = new JSDOM('<div id="rb"></div>');
global.self = dom.window; global.window = dom.window; global.document = dom.window.document;
const Spec = (await import('file://' + path.join(ROOT, 'lib/ribbon-spec.js').replace(/\\/g, '/'))).default;
dom.window.RibbonActions = new Proxy({}, { get: () => function () {} });
const Ribbon = (await import('file://' + path.join(ROOT, 'lib/ribbon.js').replace(/\\/g, '/'))).default;
const el = dom.window.document.getElementById('rb');
Ribbon.描く(el, Spec);
const タブ = [...el.querySelectorAll('.rb-tab')].map((x) => x.textContent.trim());
言う(!タブ.includes('Sheet1'),
  '★Sheet1 の タブを 描いていない★（下の シート見出しの 写り込み）', タブ.join(' / '));
言う(タブ.length === 11, '★タブ 11個★', String(タブ.length));

/* (b) の 字が 出るか＝ホーム|アドイン */
const 箱 = [...el.querySelectorAll('.rb-group')].find((g) => g.dataset.group === 'アドイン');
const 札 = 箱 ? 箱.querySelector('.rb-no') : null;
言う(!!札 && /付けません/.test(札.textContent),
  '★(b) の 箱に「付けません」と 出る★', 箱 ? 箱.textContent : '箱が 無い');
言う(!!札 && !!札.querySelector('.rb-why') && 札.querySelector('.rb-why').textContent.trim().length > 0,
  '★理由の 1行も 出る★', 札 ? 札.innerHTML : '');
言う(!!箱 && !箱.querySelector('.rb-yet'),
  '★(b) の 箱に「これから」は 出ない★');

/* ── わざと 壊して 赤に なるか ────────────────── */
if (process.argv.includes('--self-test')) {
  console.log('\n★わざと 壊して 赤に なるか★');
  const 前 = 範囲.対象外['ホーム|アドイン|アドイン'];
  delete 範囲.対象外['ホーム|アドイン|アドイン'];
  言う(範囲.付けない理由('ホーム', 'アドイン', 'アドイン') === null,
    '★理由を 消したら (b) と 見なされなくなる＝(a)に 落ちる★');
  範囲.対象外['ホーム|アドイン|アドイン'] = 前;
  言う(範囲.付けない理由('ホーム', 'アドイン', 'アドイン') === 前, '★戻した★');

  const 前C = 範囲.出さない['描画|元に戻す'];
  delete 範囲.出さない['描画|元に戻す'];
  言う(範囲.出さない理由('描画', '元に戻す') === null, '★出さないを 消したら 出る側に 戻る★');
  範囲.出さない['描画|元に戻す'] = 前C;
  言う(範囲.出さない理由('描画', '元に戻す') === 前C, '★戻した★');
}

console.log('\nribbon-scope: ' + ok + '/' + (ok + ng) + ' passed');
process.exit(ng ? 1 : 0);
