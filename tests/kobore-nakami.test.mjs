/* kobore-nakami.test.mjs — ★溢れた 先の ★中身★ を 実Excel の 紙と 突き合わせる★（2026-09-21）
 *
 *  ★★なぜ 要るか★★
 *    `tests/shiki-kansuu-kami.test.mjs` は ★1マス目だけ★ 見て います。
 *    ⇒`=ACOTH(A1:A3)` の ★2マス目が 0.549＝ で なく なっても 気付けません★
 *    （経営者1 が 2026-09-21 に そう 書きました＝★正直に 書いた 穴★）
 *    ⇒★この 門は ★溢れた 先の 全マス★ を 見ます★
 *
 *  ★★物差し★★
 *    `docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv`
 *    ・★実Excel が 打った 物★（版 16.0 build 20326）
 *    ・★柱に `実Excel（今の 道＝Formula2）` と 在る★＝★物差しに して よいという 宣言★
 *    ・★盤面も 置き場も 読む 窓も 紙の 頭に 書いて あります★
 *    ・★数は 全桁（R）★で 取り直して あります（2026-09-21）
 *
 *  ★★ここが 守る 事★★
 *    ①紙が 読めて いる（★空振りして いない★）
 *    ②★溢れた マスの 数★が 紙と 合う
 *    ③★★溢れた 先の 中身★★ が 紙と 合う（★15桁で 比べる★）
 *    ④★足りない 物を 数で 出す★（★黙って 減らない★）
 *
 *  ★★15桁で 比べる 訳★★
 *    ★実Excel が 持って いるのも 見せて いるのも 15桁までです★
 *    （`shiki-kansuu-kami.test.mjs` が 前から 同じ 手を 使って います）
 *    ⇒16桁目の 違いは ★実Excel も 見せません★
 *
 *  ★★見て いない 事★★
 *    ・★画面では ありません★（台で 押して います）
 *    ・★まだ 合って いない 物は 数だけ 出します★（★1本ずつ 名指しでは 止めません★）
 *      ＝★今 合う 数を 決め打ちで 守ります★（★減ったら 赤★）
 *
 *  使い方: node tests/kobore-nakami.test.mjs
 *          node tests/kobore-nakami.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ok   ' + n); }
  catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + e.message); }
};

const 紙道 = path.join(ROOT, 'docs/measured/golden-jitsu-excel-hanni-zenbu-2026-09-20.tsv');
const 字 = fs.existsSync(紙道) ? fs.readFileSync(紙道, 'utf8').replace(/^﻿/, '') : '';
const NL = String.fromCharCode(10);
const TB = String.fromCharCode(9);
const CR = String.fromCharCode(13);

/* ★柱から 列を 探します★（★決め打ちに しない★＝列が 増えても 動く） */
let 数列 = 2, 中列 = 3;
for (const l0 of 字.split(NL)) {
  if (!l0.startsWith('#')) continue;
  const h = l0.replace('#', '').split(TB).map((z) => String(z).trim());
  const i1 = h.findIndex((z) => z.indexOf('埋まった数') >= 0);
  const i2 = h.findIndex((z) => z.indexOf('中身') >= 0);
  if (i1 < 0 || i2 < 0) continue;
  数列 = i1; 中列 = i2;
  break;
}

/* ★材料（盤面）も 紙から 読みます★＝★道具が 決めない★ */
const 材料 = [];
const 組 = [];
for (const l0 of 字.split(NL)) {
  const l = l0.replace(CR, '');
  if (!l) continue;
  if (l.startsWith('#材料')) {
    const c = l.split(TB);
    材料.push({ マス: String(c[1] || '').trim(), 値: String(c[2] || '').trim() });
    continue;
  }
  if (l.startsWith('#')) continue;
  const c = l.split(TB);
  const 式 = String(c[0] || '').trim();
  if (!式.startsWith('=')) continue;
  const 数 = parseInt(c[数列], 10);
  const 中身 = String(c[中列] || '').trim();
  if (!isFinite(数) || !中身) continue;
  組.push({ 式, 数, 中身 });
}

console.log('');
console.log('[kobore-nakami] ★溢れた 先の 中身を 実Excel の 紙と 突き合わせる★');
console.log('      ＝ 紙 ' + 組.length + '本 ／ 材料 ' + 材料.length + '行 ／ 埋まった数=' + 数列 + '列目 ／ 中身=' + 中列 + '列目');

T('★紙が 読めて いる（空振りして いない）★', () => {
  if (!字) throw new Error('★紙が 無い★ ' + 紙道);
  if (組.length < 50) throw new Error('★' + 組.length + '本しか 読めない★');
  if (材料.length < 5) throw new Error('★材料が ' + 材料.length + '行★（盤面を 作れません）');
  /* ★名指しで 確かめます★＝★紙が 入れ替わって いないか★ */
  for (const f of ['=SEQUENCE(3)', '=ABS(A1:A3)', '=ISNUMBER(A1:A3)']) {
    if (!組.some((x) => x.式 === f)) throw new Error('★紙に ' + f + ' が 無い★');
  }
});

/* ★15桁で 揃える★（実Excel が 見せるのも 15桁） */
const 十五 = (x) => {
  const t = String(x).trim();
  /* ★★真偽の 字は 台と 紙で 書き方が 違います★★（2026-09-21 実測）
       台（JS） ＝ `false` `true`
       紙（実Excel） ＝ `FALSE` `TRUE`
     ⇒★揃えないと ★偽の 負け★に なります★（8本 そう なりました）
     ⇒★どちらが 正しいかの 話では 無く ★取り方★の 話★
       （`shiki-kansuu-kami.test.mjs` も 前から 同じ 手を 使って います） */
  if (t === 'false' || t === 'FALSE') return 'FALSE';
  if (t === 'true' || t === 'TRUE') return 'TRUE';
  const n = Number(t);
  if (!isFinite(n) || t === '') return t;
  return String(Number(n.toPrecision(15)));
};
const 揃える = (s) => String(s).split(' / ')
  .map((段) => 段.split('|').map(十五).join('|')).join(' / ');

/* ★台で 押して 溢れた 表を 取ります★ */
function 押す(式) {
  const h = H.表();
  for (const m of 材料) {
    if (!m.マス || m.値 === '') continue;
    h.打つ(m.マス, m.値);
  }
  h.打つ('BZ1', 式);
  const v = h.値('BZ1');
  /* ★溢れなら 並びを そのまま／溢れないなら 1マス★ */
  if (v && v.溢れ === true) {
    const 表 = (v.並び || []).map((段) => 段.map((x) => (x && x.値 !== undefined ? String(x.値) : String(h.字('BZ1')))));
    return { 数: 表.reduce((a, r) => a + r.length, 0), 中身: 表.map((r) => r.join('|')).join(' / ') };
  }
  return { 数: 1, 中身: String(h.字('BZ1')) };
}

let 数が合う = 0, 中身も合う = 0, 数が足りない = 0, 数が多い = 0, 投げた = 0;
const 違う = [];
for (const q of 組) {
  let 出;
  try { 出 = 押す(q.式); }
  catch (e) { 投げた++; continue; }
  if (出.数 === q.数) {
    数が合う++;
    if (揃える(出.中身) === 揃える(q.中身)) 中身も合う++;
    else 違う.push({ 式: q.式, 実: q.中身, 内: 出.中身 });
  } else if (出.数 < q.数) 数が足りない++;
  else 数が多い++;
}

console.log('      ＝ ★数が 合う★ ' + 数が合う + '本 ／ ★中身も 合う★ ' + 中身も合う + '本'
  + ' ／ 足りない ' + 数が足りない + '本 ／ 多い ' + 数が多い + '本 ／ 投げた ' + 投げた + '本');

T('★道具が 空振りして いない（1本でも 押せて いる）★', () => {
  if (数が合う + 数が足りない + 数が多い < 50) throw new Error('★押せたのが 少なすぎます★');
});

/* ★★今 合う 数を 決め打ちで 守ります★★（★減ったら 赤★＝黙って 下がらない） */
/* ★★数を 上げた 覚え★★（★黙って 増えない★＝増えたら ここを 直すまで 赤）
     ★83本★（2026-09-21 朝）＝ `見分け` `数ひとつ` `字ひとつ` まで 直した 所
     ★108本★（2026-09-21）＝ `呼ぶ` に ★値を 1マスずつ 渡す 名簿 27本★ を 足した
       ⇒この 時 `ISFORMULA` `ROW` は ★数は 合うのに 中身が 全部 `#VALUE!`★ でした
       ⇒★この 門が 見つけました★（数だけ 見る 門では 捕まりません）
     ★110本★（2026-09-21）＝ ★場所を 渡す★ 形に 分けて 直した
       `ISFORMULA` ＝ `{ 種:'マス', 番地, 打った字 }` を 1マスずつ 渡す
       `ROW` `COLUMN` ＝ ★マスごとでは なく 形の ぶんだけ★ 溢れる
     ★残り 2本★ ＝ `GROUPBY` `PIVOTBY`（★ピボットテーブル 自体が 台に 有りません★） */
T('★★中身まで 合う 数が 減って いない★★（★今 110本★）', () => {
  const 期待 = 110;
  if (中身も合う < 期待) {
    throw new Error('★' + 中身も合う + '本★（' + 期待 + '本 の はず）／★下がって います★'
      + (違う.length ? '／例 ' + 違う[0].式 + ' 実=' + String(違う[0].実).slice(0, 34) + ' 内=' + String(違う[0].内).slice(0, 34) : ''));
  }
  if (中身も合う > 期待) {
    throw new Error('★' + 中身も合う + '本に 増えました★（' + 期待 + '本 の はず）'
      + '／★増えたら ここの 数も 直して ください★（★黙って 増えない★）');
  }
});

T('★数が 多い 物は 0本（★溢れすぎて いない★）★', () => {
  if (数が多い) throw new Error('★' + 数が多い + '本 溢れすぎ★');
});

T('★投げた 物は 0本★', () => {
  if (投げた) throw new Error('★' + 投げた + '本 投げました★');
});

console.log('');
console.log('  ★まだ 合って いない★ ... 足りない ' + 数が足りない + '本 ／ 中身が 違う ' + 違う.length + '本');
console.log('  ★見て いない 事★ ... 画面では ありません（台で 押して います）');
console.log('kobore-nakami: ' + pass + ' 緑 / ' + fail + ' 赤');

if (process.argv.includes('--self-test')) {
  /* ★★わざと 壊して 赤に なるか★★＝★紙の 中身を 1本 書き換えて 押す★
       ★ここで 壊すのは 材料（紙の 写し）です★＝`lib` は 1字も 触りません */
  console.log('');
  console.log('--self-test: ★紙の 中身を 1本 変えて 押す★');
  const 的 = 組.find((x) => x.式 === '=ABS(A1:A3)');
  if (!的) { console.log('  NG   ★=ABS(A1:A3) が 紙に 無い★'); process.exit(1); }
  const 元 = 的.中身;
  /* ★表から 1行 抜いて 赤に なるかも 見ます★ */
  const 消した = 組.filter((x) => x.式 !== 的.式);
  if (消した.length !== 組.length - 1) { console.log('  NG ★表から 1行 抜けませんでした★'); process.exit(1); }
  /* ★元の 字を 書き換えます★（代入では 無く）
       ＝★紙の 数を 全部 99 に する★＝★本当に 壊して います★ */
  的.中身 = 元.replace(/[0-9]+(\.[0-9]+)?/g, '99');
  const 出 = 押す(的.式);
  const 合う = 揃える(出.中身) === 揃える(的.中身);
  的.中身 = 元;
  if (合う) { console.log('  NG   ★中身を 変えたのに 合って しまいました★'); process.exit(1); }
  console.log('  ok   ★中身を 変えると 合わなく なる（門が 中身を 見て います）★');
  process.exit(0);
}
process.exit(fail ? 1 : 0);
