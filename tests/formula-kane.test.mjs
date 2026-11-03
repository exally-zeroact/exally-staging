/* formula-kane.test.mjs — ★お金の 関数を 実Excel の 答えと 突き合わせる★（2026-09-07）
 *
 *  ★★何を 見るか★★
 *    `docs/measured/kansuu46/golden-kane-2026-09-07.tsv`（★実Excel に 打たせた 700本★）
 *    ⇒ その うち ★出すと 決めた 26個★の 分が ★1本 残らず 合う事★（2026-09-19 ... 22 → 24 → ★26★）
 *    ⇒★答えは 私が 書いていません★（Excel 16.0 build 20326 が 出した 数）
 *
 *  ★★保留の 2個も 見張る★★（2026-09-19 ... ODDL の 2個を 出しました）
 *    ODDFPRICE / ODDFYIELD / ODDLPRICE / ODDLYIELD は
 *    ★合わない 形が 残っている★ので ★出していません★。
 *    ⇒ うっかり 出してしまわないよう ★出す 名簿に 入っていない事★を ここで 見る。
 *
 *  ★★名簿は 1つだけ★★
 *    `lib/formula-kane.js` の `足した名前()` が 正本。
 *    ⇒ 繋ぐ 側（plug）が 別の 名簿を 持っていないかを 突き合わせる。
 *
 *  使い方: node tests/formula-kane.test.mjs
 *          node tests/formula-kane.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const require_ = createRequire(import.meta.url);
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));
const 金の道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');

/* ★自分が 直に 走ったかを 見る★
   ＝取り込まれた 時に argv を 見て 動くと ★別の 見張りを 乗っ取る★
   （2026-09-02 に 全社 74本で 実際に 起きた） */
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };

/* ── 式を 読む（実Excel に 打たせた 式を そのまま 使う） ── */
function 引数を割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') { 深++; 今 += ch; continue; }
    if (ch === ')') { 深--; 今 += ch; continue; }
    if (ch === ',' && 深 === 0) { 出.push(今.trim()); 今 = ''; continue; }
    今 += ch;
  }
  if (今.trim() !== '') 出.push(今.trim());
  return 出;
}
function 値にする(s) {
  s = s.trim();
  const d = s.match(/^DATE\((\d+),(\d+),(\d+)\)$/i);
  if (d) return K.日から数(+d[1], +d[2], +d[3]);
  if (/^TRUE$/i.test(s)) return true;
  if (/^FALSE$/i.test(s)) return false;
  return Number(s);
}
const 基 = (b) => (b === undefined || b === null || Number.isNaN(b) ? 0 : b);
const 呼ぶ = {
  ACCRINT: (a) => K.経過利息(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7]),
  ACCRINTM: (a) => K.満期一括の経過利息(a[0], a[1], a[2], a[3], a[4]),
  AMORDEGRC: (a) => K.仏定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  AMORLINC: (a) => K.仏定額(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  COUPDAYBS: (a) => K.前からの日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPDAYS: (a) => K.期間の日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPDAYSNC: (a) => K.次までの日数(K.数から日(a[0]), K.数から日(a[1]), a[2], 基(a[3])),
  COUPNCD: (a) => { const p = K.次の利払日(K.数から日(a[0]), K.数から日(a[1]), a[2]); return K.日から数(p.y, p.m, p.d); },
  COUPPCD: (a) => { const p = K.前の利払日(K.数から日(a[0]), K.数から日(a[1]), a[2]); return K.日から数(p.y, p.m, p.d); },
  COUPNUM: (a) => K.利払回数(K.数から日(a[0]), K.数から日(a[1]), a[2]),
  DISC: (a) => K.割引率(a[0], a[1], a[2], a[3], a[4]),
  DURATION: (a) => K.期間(a[0], a[1], a[2], a[3], a[4], a[5]),
  INTRATE: (a) => K.利率(a[0], a[1], a[2], a[3], a[4]),
  MDURATION: (a) => K.修正期間(a[0], a[1], a[2], a[3], a[4], a[5]),
  PRICE: (a) => K.価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  PRICEDISC: (a) => K.割引債の価格(a[0], a[1], a[2], a[3], a[4]),
  PRICEMAT: (a) => K.満期一括の価格(a[0], a[1], a[2], a[3], a[4], a[5]),
  RECEIVED: (a) => K.受取額(a[0], a[1], a[2], a[3], a[4]),
  VDB: (a) => K.可変定率(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  YIELD: (a) => K.利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6]),
  YIELDDISC: (a) => K.割引債の利回り(a[0], a[1], a[2], a[3], a[4]),
  YIELDMAT: (a) => K.満期一括の利回り(a[0], a[1], a[2], a[3], a[4], a[5]),
  /* ★2026-09-19 ... ODDL の 2個を 出しました★（★経営者1 の 275本で 外れ 0本★） */
  ODDLPRICE: (a) => K.最終端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], 基(a[7])),
  ODDLYIELD: (a) => K.最終端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], 基(a[7])),
  /* ★2026-09-19 ... ODDF の 2個を 出しました★（★経営者1 の 414本で 外れ 0本★＝★保留 0個★） */
  ODDFPRICE: (a) => K.初回端数の価格(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], 基(a[8])),
  ODDFYIELD: (a) => K.初回端数の利回り(a[0], a[1], a[2], a[3], a[4], a[5], a[6], a[7], 基(a[8])),
};
import { 式をほどく, 答えを見る } from './monosashi.mjs';   /* ★物差しは 1本★ */

const 二桁 = (n) => (n < 10 ? '0' : '') + n;
const 年月日 = (serial) => { const p = K.数から日(serial); return p.y + '-' + 二桁(p.m) + '-' + 二桁(p.d); };

/** ★1行 押して 合っているかを 返す★（純粋＝自己試験からも 呼べる）
 *  ★物差し（式をほどく／答えを見る）は `tests/monosashi.mjs` に 1本だけ 置いて 在ります★
 *  ＝2026-09-14。借り物を 外す 為の 自前の 皮も ★同じ物★ を 読みます。
 *  ★日付の 字に する のは ここ（群ごと）★＝物差しは 中身に 依らせない。 */
export function 押して比べる(行) {
  const ほ = 式をほどく(行.式);
  if (!ほ) return { 合: false, 訳: '式が 読めない' };
  const fn = 呼ぶ[ほ.名];
  if (!fn) return { 合: null, 訳: '出していない 関数' };
  let 出;
  try { 出 = fn(引数を割る(ほ.引数の字).map(値にする)); } catch (e) { 出 = { 誤り: 'EX' }; }
  if (ほ.字包み && typeof 出 === 'number') 出 = 年月日(出);   /* ★群ごとの 直し★ */
  return 答えを見る(出, 行);
}

const 行たち = fs.readFileSync(金の道, 'utf-8').split('\n')
  .filter((l) => l && !l.startsWith('#'))
  .map((l) => { const p = l.split('\t'); return { 名: p[0], 式: p[1], 答: p[2], 型: p[3] }; });

console.log('\n[formula-kane] お金の 関数 — 実Excel の 答えと 突き合わせ');
console.log('  実Excel の 答え … ' + 行たち.length + '本（Excel 16.0 build 20326 / UI 1041）');

const 出す = K.足した名前();
const 保留 = K.保留の名前();

T('★答えの 紙が 空でない★（材料が 無いのに 緑に しない）', () => {
  if (行たち.length < 100) throw new Error('答えが ' + 行たち.length + '本しか ない');
  if (!/build/.test(fs.readFileSync(金の道, 'utf-8').slice(0, 400))) {
    throw new Error('どの Excel で 打ったかが 書かれていない');
  }
});

T('★出す 26個は 1本 残らず 実Excel と 同じ★', () => {
  const 外れ = [];
  let 見た = 0;
  for (const 行 of 行たち) {
    if (出す.indexOf(行.名) < 0) continue;
    見た++;
    const r = 押して比べる(行);
    if (!r.合) 外れ.push(行.名 + '  ' + 行.式 + '\n        正 ' + 行.答 + ' ／ 出 ' + r.出);
  }
  if (見た < 100) throw new Error('押した 式が ' + 見た + '本しか ない（材料 不足）');
  console.log('      押した 式 … ' + 見た + '本');
  if (外れ.length) throw new Error('★' + 外れ.length + '本 違う★\n      ' + 外れ.slice(0, 5).join('\n      '));
});

T('★保留は 0個★（2026-09-19 ... ODDF の 2個を 出して 棚が 空に なった）', () => {
  for (const n of 保留) {
    if (出す.indexOf(n) >= 0) throw new Error(n + ' が 出す 名簿に 入っている');
  }
  if (保留.length !== 0) throw new Error('保留は 0個の はず（今 ' + 保留.length + '）');
});

T('★名簿は 1つだけ★（繋ぐ 側が 別の 名簿を 持っていない）', () => {
  const src = fs.readFileSync(path.join(ROOT, 'lib/formula-kane-plug.js'), 'utf-8');
  const 並 = [...src.matchAll(/^\s*'([A-Z][A-Z0-9.]*)':\s*\{\s*method:/gm)].map((m) => m[1]);
  const a = [...出す].sort().join(','), b = [...並].sort().join(',');
  if (a !== b) throw new Error('名簿が 違う\n      lib  … ' + a + '\n      plug … ' + b);
  for (const n of 保留) {
    if (並.indexOf(n) >= 0) throw new Error('保留の ' + n + ' が 繋がれている');
  }
});

T('★保留の 2個は「動かない 棚」に 訳つきで 載っている★（黙って 消えない）', () => {
  const EX = require_(path.join(ROOT, 'lib/formula-extra.js'));
  const 棚 = (EX.数える && EX.数える().足さない) || {};
  for (const n of 保留) {
    const 訳 = 棚[n];
    if (!訳) throw new Error(n + ' が 動かない 棚に 載っていない（客は #NAME? を 食らうだけ）');
    if (String(訳).length < 8) throw new Error(n + ' の 訳が 短すぎる（' + 訳 + '）');
  }
});

T('★実Excel の 赤い 値も 見ている★（数字だけ 見ていない）', () => {
  const 赤 = 行たち.filter((r) => 出す.indexOf(r.名) >= 0 && String(r.答).startsWith('#'));
  if (!赤.length) throw new Error('赤い 答えが 1本も 無い＝赤の 道を 見ていない');
  console.log('      赤い 答え … ' + 赤.length + '本（' + 赤[0].名 + ' ほか）');
});

if (自己試験) {
  console.log('\n[self-test] ★物差しが 効いているか★');
  T('わざと 1文字 変えた 答えは 赤に なる', () => {
    const 元 = 行たち.find((r) => r.名 === 'PRICE' && r.型 === 'Double');
    const 偽 = { 名: 元.名, 式: 元.式, 答: String(Number(元.答) + 1), 型: 元.型 };
    if (押して比べる(偽).合) throw new Error('違う 答えでも 緑に なった');
  });
  T('わざと 赤を 数字に すり替えたら 赤に なる', () => {
    const 元 = 行たち.find((r) => String(r.答).startsWith('#') && 出す.indexOf(r.名) >= 0);
    if (!元) throw new Error('赤い 答えが 無い');
    const 偽 = { 名: 元.名, 式: 元.式, 答: '123', 型: 'Double' };
    if (押して比べる(偽).合) throw new Error('赤なのに 数字で 緑に なった');
  });
  /* ══ ★★2026-09-19 ... ★この 自己試験の 土台が 消えました★★★ ══
       ★前★ ... `ODDFPRICE` を ★保留の 見本★として 押して いた
       ★今★ ... ★保留は 0個★（ODDF を 出した）⇒★ODDFPRICE は 押せて しまいます★
       ⇒★★「保留が 在る」前提の 試験は ★保留が 空に なった 日に 黙って 嘘に なります★★★
       ⇒★だから ★口そのもの★を 押します★＝★`呼ぶ` に 無い 名前は 「見ない」と 返るか★
       ⇒★★これなら 保留が 0個でも 100個でも 同じ 事を 見ます★★ */
  T('★`呼ぶ` に 無い 名前は 「見ない」と 返る★（★保留が 0個でも 口は 生きて いる★）', () => {
    const 無い名 = 'ZZTESTNOTREAL';
    if (呼ぶ[無い名]) throw new Error('見本に 選んだ 名前が 本当に 在る＝見本を 変えて ください');
    const r = 押して比べる({ 名: 無い名, 式: '=' + 無い名 + '(1,2)', 答: '1', 型: 'Double' });
    if (r.合 !== null) throw new Error('知らない 名前を 押してしまっている（合 ＝ ' + r.合 + '）');
    if (r.訳 !== '出していない 関数') throw new Error('訳が 違う ... ' + r.訳);
  });
  T('★保留に 名前が 戻ったら その 名前は 押さない★（★今は 0個なので 素通り★）', () => {
    for (const n of 保留) {
      const r = 押して比べる({ 名: n, 式: '=' + n + '(1,2)', 答: '1', 型: 'Double' });
      if (r.合 !== null) throw new Error('保留の ' + n + ' を 押してしまっている');
    }
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (直に走った) process.exit(fail ? 1 : 0);
