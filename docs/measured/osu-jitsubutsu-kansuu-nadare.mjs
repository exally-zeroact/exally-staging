/* osu-jitsubutsu-kansuu-nadare.mjs — ★司さんの 実物が 使う 関数の 名前を 数える★（2026-09-15）
 *
 *  ★★何の 為か★★
 *    「客に 出る 欠陥」の 紙の ★一番 上の 段★＝
 *    ★★司さんの 本に その 欠陥が 何個 当たるか★★ を 出す 為。
 *    ⇒★これが 無いと 司さんは 判じられません★
 *
 *  ★★決め（司さんの 実物・4つとも 守ります）★★
 *    ①★読むだけ＝1バイトも 書かない★（`触っていないか()` で 前後を 突き合わせ）
 *    ②★★中身を 出さない＝出してよいのは 関数の 名前と 数だけ★★
 *      （会社名・金額・人の 名前は ★1文字も 出しません★／★式そのものも 出しません★）
 *    ③★出来た 紙は repo に 入れない★＝★この 道具は 画面に 出すだけ★（紙を 書きません）
 *    ④★本体だけ★＝競合コピー／復旧／旧版は 開きません（`本の道` は 1つ）
 *
 *  ★★押し方は 書きません★★ … `osu-jitsubutsu-dodai.mjs` を 呼ぶだけ
 *    （★別の 道を 作ると 印が 抜ける★＝土台の 覚書きと 同じ）
 *
 *  ★★数え方★★
 *    ・式から ★`英字(` の 形★を 拾う（★字の 中は 先に 潰す★＝`裸に()`）
 *    ・★マスの 名前と 間違えない★ … `A1(` の ような 形は 関数では ない
 *    ・★大文字に 揃える★／★同じ 式が 何本 在るかも 数える★
 *
 *  ★★見て いない 事★★
 *    ・★名前が 在る＝使って いる★ までしか 言えません
 *      （その 式が ★どの 引数で 呼ぶか★は 見て いません）
 *    ・★お金の 21個は 「`DATE()` と 一緒に 使った 時」だけ 落ちます★
 *      ⇒★だから `DATE(` と 同じ 式に 居るか も 別に 数えます★
 *
 *  使い方: node docs/measured/osu-jitsubutsu-kansuu-nadare.mjs
 */
import { 本を開く, 触っていないか, 裸に } from './osu-jitsubutsu-dodai.mjs';

/* ★今回の 欠陥の 名簿★（★手で 並べません★＝下で 紙から 読みます） */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT } from './honban-no-michi.mjs';

const 読む = (p) => fs.readFileSync(path.join(ROOT, p), 'utf8').replace(/^﻿/, '');

/* ㋐お金＝`golden-kane-2026-09-07.tsv` に 出て 来る 関数 */
const お金 = new Set();
for (const l of 読む('docs/measured/kansuu46/golden-kane-2026-09-07.tsv').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const n = l.split('\t')[0];
  if (n) お金.add(n.toUpperCase());
}
/* ㋑D系＝`exally-formula.js` の `_jsDbFunc` が 受ける 名前（★字から 読む★） */
const D系 = new Set();
{
  const m = /\bfOrig\.match\(\/\^\(([A-Z|]+)\)/.exec(読む('exally-formula.js'));
  if (m) for (const n of m[1].split('|')) D系.add(n);
}

/* ══ ★門は 使う 前に★ ══ */
for (const [名, v] of [['お金の 名簿', お金.size], ['D系の 名簿', D系.size]]) {
  if (!v) { console.error('★★道具が 壊れて います★★ … ' + 名 + ' が ★0個★'); process.exit(1); }
}

const { wb, 前: 印 } = await 本を開く();

/* ── ★式を 集める★（★値は 見ません★＝名前だけ） ── */
const 関数拾い = /(^|[^A-Za-z0-9_.$])([A-Z][A-Z0-9_.]*)\s*\(/g;
const 数 = new Map();          /* 関数名 → 何本の 式に 出たか */
const DATEと同じ式 = new Map(); /* 関数名 → `DATE(` と 同じ 式に 出た 本数 */
let 式の本数 = 0;
let 板の数 = 0;

for (const 板名 of wb.SheetNames) {
  板の数++;
  const sh = wb.Sheets[板名];
  for (const 鍵 of Object.keys(sh)) {
    if (鍵[0] === '!') continue;
    const セル = sh[鍵];
    /* ★★`f` は ★字とは 限りません★★（2026-09-15 実測）
         ★9本が `f` を ★数★で 持って いました★（`typeof` が `number`・板 10 と 11）
         ⇒★`typeof セル.f !== 'string'` で 弾くと ★15,790本★に なり
           土台（`この板の式と値`）の ★15,799本★と ★9本 食い違う★
         ⇒★土台と 同じ `if (s.f)` に 揃える★（★数え方を 2本に しない★） */
    if (!セル || !セル.f) continue;
    式の本数++;
    const 素 = 裸に(String(セル.f)).toUpperCase();
    const 出た = new Set();
    let m;
    関数拾い.lastIndex = 0;
    while ((m = 関数拾い.exec(素)) !== null) {
      const n = m[2];
      if (/^\$?[A-Z]{1,3}\$?[0-9]{1,5}$/.test(n)) continue;   /* マスの 名前は 関数では ない */
      出た.add(n);
    }
    const DATEが居る = 出た.has('DATE');
    for (const n of 出た) {
      数.set(n, (数.get(n) || 0) + 1);
      if (DATEが居る && n !== 'DATE') DATEと同じ式.set(n, (DATEと同じ式.get(n) || 0) + 1);
    }
  }
}

/* ── ★1バイトも 触って いないか★ ── */
const 前後 = 触っていないか(印);
const 触った = 前後.字 + '（大きさ ' + 前後.大きさ + 'B）';

/* ── ★出す（★名前と 数だけ★） ── */
const 並び = [...数.entries()].sort((a, b) => b[1] - a[1]);
console.log('★読んだ 物★ 司さんの 実物 1冊 … ★板 ' + 板の数 + '枚／式 ' + 式の本数 + '本★');
console.log('   ' + 触った + '（読むだけ／1バイトも 書いて いません）');
console.log('');
console.log('★★使って いる 関数★★ ' + 並び.length + '個');
for (const [n, c] of 並び) console.log('   ' + String(c).padStart(6) + '本  ' + n);

const 当たる = (集, 札) => {
  const 重 = 並び.filter(([n]) => 集.has(n));
  console.log('');
  console.log('★★' + 札 + '★★ 名簿 ' + 集.size + '個');
  console.log('   ★この 本で 使って いる★ ' + 重.length + '個'
    + (重.length ? ' … ' + 重.map(([n, c]) => n + '(' + c + '本)').join(' ') : ' ⇒★★0個★★'));
  return 重;
};
const 金重 = 当たる(お金, 'お金の 関数');
const D重 = 当たる(D系, 'D系（`_jsDbFunc`）');

console.log('');
console.log('★★`DATE(` と 同じ 式に 居る お金の 関数★★（★落ちる のは この 形だけ★）');
const 同居 = 金重.filter(([n]) => DATEと同じ式.has(n));
console.log('   ' + (同居.length ? 同居.map(([n]) => n + '(' + DATEと同じ式.get(n) + '本)').join(' ') : '★★0個★★'));
console.log('');
console.log('★この 本に `DATE(` が 在る 式★ ' + (数.get('DATE') || 0) + '本');
