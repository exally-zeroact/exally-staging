/* ribbon-match.mjs — ★Excelのリボン288部品と うちの操作を 名前で 突き合わせる★ 2026-08-29
 *
 *  ★なぜ在るか（司さん 2026-08-29）★
 *    「リボンは ★配置なども真似しろ★ って前から言うてるよな」
 *    「★未測定★で報告を止めるな（サボるなや）」
 *    ⇒ 分母（288）だけ書いて「うちは未測定」で 終わらせない。
 *      ★名前が そのまま 一致する数★を 機械で 数えて 出す。
 *
 *  ★この道具が 出すのは「名前が一致した数」であって「同じ働きをする数」ではない★
 *    ・名前が違うだけで 在る物は 数に入らない（＝★少なめに 出る★）
 *    ・名前が同じでも 中身が 浅い物は 数に入ってしまう（＝★多めに 出る★）
 *    ⇒ ★どちらにも 転ぶ★ので、必ず ★合わなかった一覧★も 一緒に 出す。
 *      「◯/288 出来ている」とは 言わない。★「名前が一致 ◯/288」と 言う★。
 *
 *  使い方:
 *    node scripts/ribbon-match.mjs              … 数と 内訳
 *    node scripts/ribbon-match.mjs --nokori     … 合っていない物を 全部 出す
 *    node scripts/ribbon-match.mjs --json       … 機械で 読む形
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { JSDOM } from 'jsdom';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/* ── ① Excel側（正本＝機械で取った物）───────────────────── */
const TSV = path.join(ROOT, 'docs/excel-ribbon-flat.tsv');
if (!fs.existsSync(TSV)) {
  console.error('★正本が 無い★: docs/excel-ribbon-flat.tsv（tools/ribbon-dump.ps1 で 作る）');
  process.exit(1);
}
const 行 = fs.readFileSync(TSV, 'utf8').replace(/^﻿/, '').split(/\r?\n/).filter((l) => l.trim());
const 三つ組 = Array.from(new Set(行)).map((l) => l.split('\t')).filter((c) => c.length >= 3);
const リボン = 三つ組.map(([タブ, 群, 部品]) => ({ タブ, 群, 部品: 部品.trim() }));

/* ── ② うち側（本物の book.html を 載せて DOM から 拾う）──── */
const html = fs.readFileSync(path.join(ROOT, 'book.html'), 'utf8');
const dom = new JSDOM(html, { runScripts: 'outside-only' });
const doc = dom.window.document;
/* ★ログインの窓は 外す★（Excel には 無い物＝混ぜると「近づいた」に見えて 嘘になる。
   tests/excel-parity.test.mjs と 同じ決まり） */
const 操作 = [...doc.querySelectorAll('button, select, input, [onclick], [title], option')]
  .filter((el) => !el.closest('#loginOv'))
  .flatMap((el) => [
    (el.textContent || '').trim(),
    el.getAttribute('title') || '',
    el.getAttribute('aria-label') || '',
    el.getAttribute('placeholder') || '',
  ])
  .map((s) => String(s).replace(/\s+/g, ' ').trim())
  .filter(Boolean);

/* ── ③ 名前を そろえる（★中黒・かっこ・記号・空白を 落とす★）─ */
function ならす(s) {
  return String(s)
    .replace(/[（(][^）)]*[）)]/g, '')      // （…）の中は 落とす
    .replace(/\.\.\.$|…$/, '')              // 末尾の …
    .replace(/[\s・･/／,、。:：]/g, '')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0))
    .toLowerCase();
}
const うちの名 = new Set(操作.map(ならす).filter((s) => s.length >= 2));

/* ── ④ 突き合わせ ───────────────────────────────── */
const 合った = [], 合わない = [];
for (const r of リボン) {
  const k = ならす(r.部品);
  if (k.length >= 2 && うちの名.has(k)) 合った.push(r);
  else 合わない.push(r);
}
const タブごと = {};
for (const r of リボン) (タブごと[r.タブ] ||= { 全: 0, 一致: 0 }).全++;
for (const r of 合った) タブごと[r.タブ].一致++;

/* ── ⑤ ★手で結んだ対応表★ ─────────────────────────────
 *  名前だけの照合は ★少なめに 出る★（うちの物は JSが 動いてから 作られる物が 多く、
 *  静かな HTML には 出てこない。実測：並べ替え・絞り込み・検索・条件付き書式は
 *  静的な DOM に 1個も 無かった＝2026-08-29）。
 *  そこで ★作った物を 手で リボンの三つ組に 結ぶ★。
 *  ★手で書いた物は 嘘をつく★ので、結んだ先が ★正本の中に 実在するか★を 機械で 確かめる。
 *  実在しない三つ組を 書いたら ★赤にする★（Excelの版が変わって 名前が変わった時も 気づける）。 */
const 手で結んだ = [
  ['ホーム', '編集', 'オート SUM'],
  ['数式', '関数ライブラリ', 'オート SUM'],
  ['ホーム', '編集', '並べ替えとフィルター'],
  ['データ', '並べ替えとフィルター', '昇順'],
  ['データ', '並べ替えとフィルター', '降順'],
  ['データ', '並べ替えとフィルター', 'フィルター'],
  ['ホーム', '編集', '検索と選択'],
  ['ホーム', 'スタイル', '条件付き書式'],
  ['表示', 'ウィンドウ', 'ウィンドウ枠の固定'],
  ['データ', 'データ ツール', 'データの入力規則...'],
];
const 正本の集合 = new Set(リボン.map((r) => r.タブ + '' + r.群 + '' + r.部品));
const 実在しない = 手で結んだ.filter((t) => !正本の集合.has(t.join('')));
const 結べた = 手で結んだ.filter((t) => 正本の集合.has(t.join('')));
/* 名前一致と 重ならない物だけ 足す */
const 一致の鍵 = new Set(合った.map((r) => r.タブ + '' + r.群 + '' + r.部品));
const 手で足した = 結べた.filter((t) => !一致の鍵.has(t.join('')));

if (process.argv.includes('--json')) {
  console.log(JSON.stringify({ 全: リボン.length, 一致: 合った.length, タブごと, 合わない }, null, 1));
  process.exit(0);
}

console.log('[ribbon-match] ★Excelのリボン と うちの操作を 名前で 突き合わせた★');
console.log('  Excel側 … ' + リボン.length + '部品（重なりを 除いた三つ組）');
console.log('  うち側 … ' + うちの名.size + '個の 名前（book.html の DOM から 拾った）');
console.log('');
console.log('  ★名前が 一致 … ' + 合った.length + '/' + リボン.length
  + '（' + (合った.length / リボン.length * 100).toFixed(1) + '%）★');
console.log('');
console.log('  ── タブごと ──');
for (const t of Object.keys(タブごと)) {
  const x = タブごと[t];
  console.log('    ' + t.padEnd(14, '　') + ' ' + String(x.一致).padStart(3) + ' / ' + String(x.全).padStart(3));
}
console.log('');
console.log('  ★この数は「同じ働きをする数」ではない★');
console.log('    ・名前が違うだけで 在る物は 入らない（少なめに 出る）');
console.log('    ・名前が同じでも 中身が 浅い物は 入ってしまう（多めに 出る）');
console.log('    ⇒ ★「名前が一致 ' + 合った.length + '/' + リボン.length + '」としか 言わない★');
console.log('');
console.log('  ── ★手で結んだ分（作った物を リボンの三つ組に 結んだ）── ');
console.log('    結んだ … ' + 手で結んだ.length + '個 ／ 正本に 実在した … ' + 結べた.length + '個'
  + ' ／ ★実在しない … ' + 実在しない.length + '個★');
for (const t of 実在しない) console.log('       ★正本に 無い★ ' + t.join(' / ') + '（Excelの版で 名前が 変わった可能性）');
console.log('    名前一致と 重ならない分 … +' + 手で足した.length + '個');
console.log('');
console.log('  ★★合わせて ' + (合った.length + 手で足した.length) + '/' + リボン.length
  + '（' + ((合った.length + 手で足した.length) / リボン.length * 100).toFixed(1) + '%）★★'
  + ' … ★これでも 下限★（手で結べていない物が まだ 在る）');
if (実在しない.length) process.exitCode = 1;
if (process.argv.includes('--nokori')) {
  console.log('\n  ── 合っていない ' + 合わない.length + '個 ──');
  let 前 = '';
  for (const r of 合わない) {
    if (r.タブ + r.群 !== 前) { console.log('    【' + r.タブ + ' / ' + r.群 + '】'); 前 = r.タブ + r.群; }
    console.log('       ・' + r.部品);
  }
}
