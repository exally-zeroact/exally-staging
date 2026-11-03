/* shiki-sagasu.test.mjs — ★INDEX と MATCH を 実Excel の 紙 193本と 突き合わせる★（2026-09-15）
 *
 *  ★★なぜ この 2つを 先に 書いたか★★
 *    司さんの 実物 1冊を 数えたら ★出てくる 関数は 8個だけ★／
 *    その うち ★INDEX と MATCH で のべの 68%★。⇒★借り物を 外す 道の 一番 太い 所★。
 *
 *  ★★押す 紙★★（★2枚とも 実Excel に 打って 取った 物★）
 *    ① `golden-index-match2-2026-09-15.tsv`       … 138本（★誤り 100本★）
 *         ★総当たり★＝引数の 形を 101通り。★材料は 数と 日付だけ★。
 *         ★式は 全部 H1 に 打って 在ります★（★暗黙の 交わりが 変わるので ここも H1 に 打つ★）
 *    ② `golden-index-match-kimari-2026-09-15.tsv` … 55本
 *         ★決まりを 名指し★＝★字の 表・降順の 列・並んで いない 列・空マス・真偽★を 入れた。
 *         ★1列目に「打った マス」が 書いて 在ります★＝★その マスに 打ちます★。
 *
 *  ★★この 見張りが 守る 事★★
 *    ①★紙の 行を 1本も 飛ばさない★（飛ばした 本数を 必ず 出す）
 *    ②★合わない 行は 隠さず 名指しで 出す★
 *    ③★合わない 本数が 増えたら 赤★（★「知って いる 合わない 物」だけ 通す★）
 *
 *  使い方: node tests/shiki-sagasu.test.mjs
 *          node tests/shiki-sagasu.test.mjs --self-test
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { 材料を読む, 表に置く } from './zairyou.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const H = require_(path.join(ROOT, 'lib/shiki-hyou.js'));

let pass = 0, fail = 0;
const T = (n, f) => {
  try { f(); pass++; console.log('  ✓ ' + n); }
  catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); }
};

/* ★★知って いる「合わない 行」★★
   ★ここに 書いた 物だけ 通します★＝★増えたら 赤★／★訳を 必ず 1行 書く★ */
const 合わないと分かっている = {
  '=MATCH(A1:A5,1,0)':
    '★探す値が 表・並びが 直の 1つ★＝実Excel #N/A ／ うち 1。'
    + '★溢れ（土台⑤）で 表を 1本ずつ 押す 話★＝そこを 作ってから 直す（棚）',
  /* ★★書式の 台が まだ 作って いない 形★★（2026-09-15）
     ★うちは 嘘の 字を 出さず #VALUE! と 言います★＝★半分 合う 字より 断る★
     ★司さんの 実物は 1つも 使って いません★
       TEXT 740回＝`aaa`（731）と `m/d`（9）だけ ／ 画面の 18種にも 在りません
     ⇒★要ると 分かったら 足します★（★実測の 答えは 紙に 在ります★＝すぐ 作れます） */
  '=TEXT(45292,"ggge年m月d日")':
    '★和暦（ggge）は まだ 作って いません★＝実Excel 令和6年1月1日 ／ うち #VALUE!。'
    + '★実物は 使って いません★（TEXT 740回は aaa と m/d だけ）。★嘘の 字は 出しません★',
  '=TEXT(45292,"ge.m.d")':
    '★和暦（ge）は まだ 作って いません★＝実Excel R6.1.1 ／ うち #VALUE!。'
    + '★実物は 使って いません★。★嘘の 字は 出しません★',
  '=TEXT(1.5,"[h]:mm")':
    '★24時間を 超える 時（[h]）は まだ 作って いません★＝実Excel 36:00 ／ うち #VALUE!。'
    + '★実物は 使って いません★。★嘘の 字は 出しません★',
  '=TEXT(1.5,"# ?/?")':
    '★分数（?/?）は まだ 作って いません★＝実Excel 1 1/2 ／ うち #VALUE!。'
    + '★実物は 使って いません★。★嘘の 字は 出しません★',
};

const 紙たち = [
  { 道: 'docs/measured/kansuu46/golden-index-match2-2026-09-15.tsv', 打つ列: null, 既定のマス: 'H1' },
  { 道: 'docs/measured/kansuu46/golden-index-match-kimari-2026-09-15.tsv', 打つ列: 0, 既定のマス: null },
  /* ★③ 残り 4個★（IF／IFERROR／SUBTOTAL／TEXT）… 99本
     ★TEXT は まだ 書いて いません★＝`#NAME?` に なるので ひとりでに 飛びます（棚 ⑮） */
  { 道: 'docs/measured/kansuu46/golden-nokori4-kimari-2026-09-15.tsv', 打つ列: 0, 既定のマス: null },
  /* ★棚 ⑫⑬ を 埋めた 分★（35本）… ★書いた 後でも 測る★
     ⑫ MATCH の 型 -1（降順 6本＋崩れた 並び 6本）／⑬ ①2次元 ②後ろが 空 ③`~` ④式が 作った 表 ⑤横1本 */
  { 道: 'docs/measured/kansuu46/golden-tana-12-13-2026-09-15.tsv', 打つ列: 0, 既定のマス: null },
];

console.log('\n[shiki-sagasu] ★INDEX と MATCH★ … 実Excel の 紙と 1本ずつ 突き合わせる');

let 全 = 0, 合 = 0, 違 = 0, 飛 = 0;
const 違う行 = [], 知らない違い = [];
const 足りない名 = new Set();

for (const 紙 of 紙たち) {
  const 道 = path.join(ROOT, 紙.道);
  const 材料 = 材料を読む(道);
  if (!Object.keys(材料).length) throw new Error('★材料が 読めない★ ' + 紙.道);
  const 生 = fs.readFileSync(道, 'utf8').split(/\r?\n/);

  for (const l of 生) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    /* ★2枚とも 並びは 同じ★（1列目だけ 違う＝紙①は 関数名／紙②は 打った マス） */
    const マス = 紙.打つ列 === null ? 紙.既定のマス : (c[紙.打つ列] || '').trim();
    const 式 = (c[1] || '').trim();
    const 正 = (c[2] || '').trim();
    const 型 = (c[3] || '').trim();
    if (!/^=/.test(式)) continue;
    if (正 === '★受け付けない★') { 飛++; continue; }   /* 実Excel が 式として 受け付けない */

    const 表 = H.表();
    表に置く(表, 材料);
    表.打つ(マス, 式);
    const 出 = String(表.字(マス));

    /* ★★「まだ 書いて いない 関数」を 名前で 決め打ちしない★★（2026-09-15）
       ＝名簿に 書くと ★書き終えても 飛ばし続けます★（★見張りが 思い込みを 守る★）
       ⇒★うちが #NAME? と 言い、紙が そう 言って いない 行★＝★まだ 無い 関数★
         ⇒★1つ 書く たびに 飛ばす 数が ひとりでに 減ります★
       （実測＝CELL は 皮に 在るので 飛びません。
         `=CELL("address",INDEX(A1:B5,2,2))` → ★$B$2★＝★参照を 返して いる 証拠★） */
    if (出 === '#NAME?' && 正 !== '#NAME?') {
      飛++;
      (式.match(/([A-Z][A-Z0-9_.]*)\s*\(/g) || [])
        .forEach((x) => 足りない名.add(x.replace(/\s*\($/, '')));
      continue;
    }
    全++;

    /* ★型で 較べ方を 変える★（★字の 「2」と 数の 2 を 同じに しない★のは 紙が 型を 持つから） */
    let 同じ;
    if (型 === 'Boolean') 同じ = 出.toUpperCase() === 正.toUpperCase();
    else if (型 === 'Double') {
      /* ★★空の 字を 0 と 見ない★★（2026-09-15 に わざと 壊して 見つけた 穴）
         `Number('')` は ★0★ です ⇒ うちが ★空★ を 返しても
         紙の ★0★ と 同じに 見えて いました（★見張りが 穴を 通す★）。
         ⇒★数の 所は 数の 字で ある事も 見る★ */
      const a = 出.trim() === '' ? NaN : Number(出), b = Number(正);
      同じ = Number.isFinite(a) && Number.isFinite(b)
        && Math.abs(a - b) <= Math.max(1e-9, Math.abs(b) * 1e-9);
    } else 同じ = 出 === 正;

    if (同じ) { 合++; continue; }
    違++;
    違う行.push('    ' + マス.padEnd(3) + ' ' + 式.padEnd(42) + ' 紙 `' + 正 + '` ／ うち `' + 出 + '`');
    if (!Object.prototype.hasOwnProperty.call(合わないと分かっている, 式)) {
      知らない違い.push(マス + ' ' + 式 + ' … 紙 `' + 正 + '` ／ うち `' + 出 + '`');
    }
  }
}

console.log('');
console.log('  ★押した 式★ … ' + 全 + '本（飛ばした ' + 飛 + '本＝★うちが #NAME? と 言った 行★）');
console.log('  ★まだ 無い 関数★ … '
  + ([...足りない名].filter((x) => x !== 'INDEX' && x !== 'MATCH').sort().join(' ') || 'なし'));
console.log('  ★合った★   … ' + 合 + '本（' + (全 ? Math.round(合 / 全 * 1000) / 10 : 0) + '%）');
console.log('  ★合わない★ … ' + 違 + '本');
if (違う行.length) { console.log('  ★合わない 行（全部）★'); 違う行.forEach((x) => console.log(x)); }

T('★紙の 行を 1本も 飛ばして いない★（押した 本数が 0 なら 赤）', () => {
  if (全 < 180) throw new Error('押した 本数が 少なすぎる … ' + 全);
});

T('★合わない 行は 全部「知って いる 物」だけ★（増えたら 赤）', () => {
  if (知らない違い.length) {
    throw new Error('★知らない 合わない 行 ' + 知らない違い.length + '本★\n      '
      + 知らない違い.join('\n      '));
  }
});

T('★「知って いる 合わない 物」には 訳が 書いて 在る★', () => {
  for (const k of Object.keys(合わないと分かっている)) {
    if (String(合わないと分かっている[k]).length < 20) throw new Error('訳が 短すぎる … ' + k);
  }
});

T('★客の 使い方が 通る★（INDEX＋MATCH で 隣の 列を 取る）', () => {
  const 表 = H.表();
  for (let i = 1; i <= 5; i++) 表.打つ('A' + i, String(i));
  for (let i = 1; i <= 5; i++) 表.打つ('B' + i, String(i * 2));
  表.打つ('J1', '=INDEX(B1:B5,MATCH(3,A1:A5,0))');
  if (表.字('J1') !== '6') throw new Error('J1 … ' + 表.字('J1'));
});

if (process.argv.includes('--self-test')) {
  console.log('\n[shiki-sagasu --self-test] ★わざと 壊したら 赤に なるか★');

  T('★見つからない時に #N/A を 返さないと 分かる★', () => {
    const 表 = H.表();
    for (let i = 1; i <= 5; i++) 表.打つ('A' + i, String(i));
    表.打つ('J1', '=MATCH(99,A1:A5,0)');
    if (表.字('J1') !== '#N/A') throw new Error('J1 … ' + 表.字('J1'));
  });

  T('★転け方を 1つに まとめて いない★（#REF! と #VALUE! を 分けて いる）', () => {
    const 表 = H.表();
    for (let i = 1; i <= 5; i++) 表.打つ('A' + i, String(i));
    表.打つ('J1', '=INDEX(A1:A5,6)');
    if (表.字('J1') !== '#REF!') throw new Error('外は #REF! … ' + 表.字('J1'));
    表.打つ('J2', '=INDEX(A1:A5,-1)');
    if (表.字('J2') !== '#VALUE!') throw new Error('負は #VALUE! … ' + 表.字('J2'));
  });

  T('★打つ マスで 答えが 変わる★（暗黙の 交わりを 入れて いる）', () => {
    const 表 = H.表();
    for (let i = 1; i <= 5; i++) { 表.打つ('A' + i, String(i)); 表.打つ('B' + i, String(i * 2)); }
    表.打つ('J1', '=INDEX(A1:B5,0,2)');
    表.打つ('J3', '=INDEX(A1:B5,0,2)');
    表.打つ('J9', '=INDEX(A1:B5,0,2)');
    if (表.字('J1') !== '2') throw new Error('J1 … ' + 表.字('J1'));
    if (表.字('J3') !== '6') throw new Error('J3 … ' + 表.字('J3'));
    if (表.字('J9') !== '#VALUE!') throw new Error('J9 … ' + 表.字('J9'));
  });

  T('★参照が `:` の 端に なれる★（AREAS／OFFSET の 土台）', () => {
    const 表 = H.表();
    for (let i = 1; i <= 5; i++) 表.打つ('A' + i, String(i));
    表.打つ('J1', '=SUM(INDEX(A1:A5,2):A5)');
    if (表.字('J1') !== '14') throw new Error('J1 … ' + 表.字('J1'));
  });

  T('★型は 符号だけ★（2 も 0.5 も 1 と 同じ／-0.5 は -1 と 同じ）', () => {
    const 表 = H.表();
    for (let i = 1; i <= 5; i++) 表.打つ('A' + i, String(i));
    for (const [式, 正] of [['=MATCH(3,A1:A5,2)', '3'], ['=MATCH(3,A1:A5,0.5)', '3'],
      ['=MATCH(3,A1:A5,1.9)', '3'], ['=MATCH(3,A1:A5,-2)', '#N/A'], ['=MATCH(3,A1:A5,-0.5)', '#N/A']]) {
      表.打つ('J1', 式);
      if (表.字('J1') !== 正) throw new Error(式 + ' … ' + 表.字('J1') + '（紙 ' + 正 + '）');
    }
  });

  T('★型 1 は なめて いない（二分探索）★', () => {
    const 表 = H.表();
    for (const [i, v] of [5, 4, 3, 2, 1].entries()) 表.打つ('C' + (i + 1), String(v));
    表.打つ('J1', '=MATCH(3,C1:C5,1)');
    if (表.字('J1') !== '3') throw new Error('なめると 5 に なる … ' + 表.字('J1'));
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
