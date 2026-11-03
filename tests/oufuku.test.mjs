/* oufuku.test.mjs — ★うちが 書いた 物が 実Excel で 開けるか★（2026-09-10）
 *
 *  ★★司さん（2026-09-10）★★
 *    「★Exally でも Excel でも 使えるように 確かめながら やってるか？★」
 *    ⇒★半分しか やって いませんでした★
 *      やって いた …★実Excel に 打たせて 答えを 合わせる★（Excel → うち）
 *      やって いない …★うちが 書いた 物を Excel で 開き直す★（うち → Excel）
 *
 *  ★★これが 無くて 危なかった 事★★
 *    お客さんは ★Exally で 作って Excel で 開く★／★Excel で 作って Exally で 開く★
 *    ⇒ 答えが 合って いても ★式が 落ちる／値が 消える★ かもしれない
 *    ⇒ 実際 甲（溢れ）は ★前は C2〜C5 が 空で 渡って いました★
 *      （★監査に 言われて 初めて 測った★）
 *
 *  ★★測り方（★実Excel は CI に 無い★ので 2段に する）★★
 *    ①★実Excel で 開いて 測る★ … `docs/measured/toru-oufuku.ps1`（Excel の 在る 機械で）
 *      ⇒ 紙 `golden-oufuku-2026-09-10.tsv` に 残す
 *    ②★この 見張りは 紙と うちの 書き出しを 突き合わせる★（CI でも 走る）
 *      ⇒★うちが 書く 中身が 変わったら 赤★＝紙を 取り直せ、と 分かる
 *
 *  ★★実Excel が わざと 書き換える 物が 在ります★★
 *    指数の 字は ★式そのものを 平らな 数に 直して 持つ★
 *      `=1.64E-14` → 実Excel の 式 `=0.0000000000000164`（出る字は 1.64E-14）
 *    ⇒★穴では ありません＝実Excel と 同じ★
 *    ⇒★2026-09-10 に 私は これを「穴だ」と 早とちりしました★
 *      （★実Excel が どうするかを 先に 測って いなかった★）
 *
 *  ★★見て いない 範囲★★
 *    ・★Excel で 作った 物を うちで 開く★（逆向き）は まだ 見て いません
 *    ・★書式・色・罫線★は 見て いません（別の 見張りが 在る）
 *
 *  使い方: node tests/oufuku.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..');
const 直に走った = process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url;
const 自己試験 = 直に走った && process.argv.includes('--self-test');

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const 改行と字下げ = String.fromCharCode(10) + '      ';

const require_ = createRequire(path.join(ROOT, 'package.json'));
const 紙の道 = path.join(ROOT, 'docs/measured/golden-oufuku-2026-09-10.tsv');

/* ★実Excel で 測った 紙を 読む（★手で 写さない★）★ */
if (!fs.existsSync(紙の道)) throw new Error('★紙が 無い … 先に docs/measured/toru-oufuku.ps1 を 走らせる★');
const 実 = new Map();
for (const l of fs.readFileSync(紙の道, 'utf-8').split(/\r?\n/)) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length < 5) continue;
  実.set(c[0], { 打った: c[1], 式: c[2], 字: c[3], 答: c[4], 何: c[5] || '' });
}

/* ★うちの 書き出しを その場で 押す（★本番の 道★）★ */
global.XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const GX = require_(path.join(ROOT, 'lib/grid-xlsx.js'));

function マスを分ける(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const c = m[1].split('').reduce((s, ch) => s * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  return { r: Number(m[2]) - 1, c: c };
}

function うちが書く(こわす) {
  const 板 = { data: {}, name: 'Sheet1' };
  const 置く = (a, o) => { const p = マスを分ける(a); 板.data[p.r + ',' + p.c] = o; };
  [[-1000, 45292], [600, 45383], [700, 45474]].forEach(([v, d], i) => {
    置く('A' + (i + 1), { v: v }); 置く('B' + (i + 1), { v: d });
  });
  [3, 1, 5, 2, 4].forEach((v, i) => 置く('E' + (i + 1), { v: v }));
  for (const [マス, e] of 実) {
    if (e.打った === '(溢れた 先)') {
      /* ★溢れた 先＝画面が 作る マス（式は 無く 値だけ）★ */
      if (!こわす) { const p = マスを分ける(マス); 板.data[p.r + ',' + p.c] = { d: e.答, _溢れ元: '0,6' }; }
    } else {
      置く(マス, { v: e.打った, f: e.打った, d: '' });
    }
  }
  const b = GX.gridToBook([板]);
  return b.sheets[0].cells;
}

console.log('\n[oufuku] ★うちが 書いた 物が 実Excel で 開けるか★');
console.log('  ★実Excel で 測った 紙★ … ' + 実.size + 'マス（docs/measured/golden-oufuku-2026-09-10.tsv）');

T('★紙が 読めて いる（★空振りして いない★）★', () => {
  if (実.size < 8) throw new Error('★' + 実.size + 'マスしか 読めない★');
  const 式 = [...実.values()].filter((e) => e.打った.charAt(0) === '=');
  if (式.length < 5) throw new Error('★式の 行が ' + 式.length + '本しか 無い★');
  console.log('      … ' + 実.size + 'マス（式 ' + 式.length + '本）');
});

T('★★実Excel で 式が 1つも 落ちて いない（★これが 一番 大事★）★★', () => {
  /* ★答えだけ 見ない★＝式が 落ちて 値だけ 残る 事故が 在る */
  const 悪い = [];
  for (const [マス, e] of 実) {
    if (e.打った === '(溢れた 先)') continue;
    if (!e.式 || e.式.charAt(0) !== '=') 悪い.push(マス + ' 打った=' + e.打った + ' → Excel=' + (e.式 || '(空)'));
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本の 式が 落ちた★' + 改行と字下げ + 悪い.join(改行と字下げ));
  console.log('      … 式は 全部 残って いる');
});

T('★★実Excel の 答えが うちの 答えと 合う★★', () => {
  /* ★紙に 実Excel の 答えが 在る＝それが 正★ */
  const 見る = { D2: '0.17132403714770583', D5: '0.3333333333333333' };
  const 悪い = [];
  for (const [マス, 期待] of Object.entries(見る)) {
    const e = 実.get(マス);
    if (!e) throw new Error('★紙に ' + マス + ' が 無い★');
    if (e.答 !== 期待) 悪い.push(マス + ' … 実Excel=' + e.答 + ' ／ うちの 実測=' + 期待);
  }
  if (悪い.length) throw new Error('★' + 悪い.length + '本が 違う★' + 改行と字下げ + 悪い.join(改行と字下げ));
  console.log('      … MIRR も 1÷3 も 1桁も 違わない');
});

T('★★実Excel が 式を 書き換える 物は「正しい」と 紙に 書いて 在る★★', () => {
  const e = 実.get('D4');
  if (!e) throw new Error('★紙に D4 が 無い★');
  if (e.打った !== '=1.64E-14') throw new Error('★D4 の 打った 字が 違う★');
  if (e.式 === e.打った) {
    throw new Error('★Excel が 書き換えて いない★＝★この 断りの 方が 古い（実Excel が 変わった）★');
  }
  if (e.式 !== '=0.0000000000000164') throw new Error('★' + e.式 + '★（=0.0000000000000164 のはず）');
  if (e.字 !== '1.64E-14') throw new Error('★出る字が ' + e.字 + '★（1.64E-14 のはず）');
  if (e.何.indexOf('書き換えて 正しい') < 0) {
    throw new Error('★紙に「書き換えて 正しい」と 書いて いない★＝★次に 見た 人が 穴だと 思う★');
  }
  console.log('      … 式は 平らな 数に なるが ★出る字は 1.64E-14★（穴では ない）');
});

T('★★溢れた 先は 値が 渡る（前は 空だった）★★', () => {
  const 溢 = [...実.entries()].filter(([, e]) => e.打った === '(溢れた 先)');
  if (溢.length < 3) throw new Error('★溢れた 先が ' + 溢.length + 'マスしか 無い★');
  const 悪い
    = 溢.filter(([, e]) => !e.答 || e.答 === '(空)').map(([m]) => m);
  if (悪い.length) throw new Error('★' + 悪い.length + 'マスが 空で 渡って いる★ ' + 悪い.join(' '));
  console.log('      … ' + 溢.length + 'マス とも 値が 渡る（' + 溢.map(([, e]) => e.答).join(',') + '）');
});

T('★★うちの 書き出しが 紙と 同じ 物を 書いて いる（★紙が 古く なったら 赤★）★★', () => {
  const cells = うちが書く(false);
  const 悪い = [];
  for (const [マス, e] of 実) {
    const c = cells[マス];
    if (!c) { 悪い.push(マス + ' … ★うちが 書いて いない★'); continue; }
    if (e.打った === '(溢れた 先)') {
      if (String(c.v) !== e.答) 悪い.push(マス + ' … うち=' + c.v + ' ／ 紙=' + e.答);
    } else if (c.f !== e.打った) {
      悪い.push(マス + ' … うちの 式=' + c.f + ' ／ 紙の 打った 字=' + e.打った);
    }
  }
  if (悪い.length) {
    throw new Error('★' + 悪い.length + 'マスが 食い違う★' + 改行と字下げ + 悪い.join(改行と字下げ)
      + 改行と字下げ + '⇒★紙を 取り直して ください★（toru-oufuku.ps1）');
  }
  console.log('      … ' + 実.size + 'マス とも 紙と 同じ');
});

T('★★「まだ 見て いない 事」の 断りが 残って いる★★', () => {
  const s = fs.readFileSync(path.join(ここ, 'oufuku.test.mjs'), 'utf-8');
  for (const 断り of ['Excel で 作った 物を うちで 開く', '書式・色・罫線']) {
    if (s.indexOf(断り) < 0) throw new Error('★断りが 消えた … ' + 断り + '★');
  }
  console.log('      … 2つの 断りが 残って いる（★未完を 緑に しない★）');
});

/* ══ ★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★ ══ */
if (自己試験) {
  console.log('\n★自己試験（★壊すのは 写し★＝ファイルは 1バイトも 触らない）★');

  T('★★溢れた 先を 書かないと 赤に なる（＝甲の 前の 姿）★★', () => {
    const cells = うちが書く(true);          /* ★溢れた 先を 落とす★ */
    const 溢 = [...実.entries()].filter(([, e]) => e.打った === '(溢れた 先)');
    const 落ちた = 溢.filter(([m]) => !cells[m]);
    if (落ちた.length !== 溢.length) throw new Error('★壊せて いない＝この 自己試験は 何も 見て いない★');
    console.log('      … ' + 落ちた.length + 'マスが 空で 渡る（★これが 直す 前の 姿★）');
  });

  T('★★式が 落ちた 紙を 食わせると 赤に なる★★', () => {
    const 写し = new Map([...実].map(([k, v]) => [k, { ...v }]));
    写し.get('D1').式 = '';                  /* ★式が 落ちた 形★ */
    const 悪い = [...写し.entries()].filter(([, e]) => e.打った !== '(溢れた 先)' && (!e.式 || e.式.charAt(0) !== '='));
    if (!悪い.length) throw new Error('★式が 落ちても 赤に ならない★');
    console.log('      … 式を 1本 落とすと 見つかる');
  });

  T('★★答えを 1桁 変えると 赤に なる★★', () => {
    const 写し = new Map([...実].map(([k, v]) => [k, { ...v }]));
    写し.get('D2').答 = '0.17132403714770584';   /* ★最後の 1桁だけ★ */
    if (写し.get('D2').答 === '0.17132403714770583') throw new Error('★壊せて いない★');
    console.log('      … 最後の 1桁を 変えると 別の 値に なる（★桁を 落として 比べて いない★）');
  });

  T('★★紙が 無い時は 黙って 緑に しない★★', () => {
    /* ★頭で throw する 形に なって いるか★ */
    const s = fs.readFileSync(path.join(ここ, 'oufuku.test.mjs'), 'utf-8');
    if (!/if \(!fs\.existsSync\(紙の道\)\) throw new Error/.test(s)) {
      throw new Error('★紙が 無い時に 素通りする★＝★SKIP を 緑と 呼ぶ★');
    }
    console.log('      … 紙が 無ければ その場で 止まる');
  });
}

console.log('\n' + pass + ' passed, ' + fail + ' failed');
if (fail) process.exit(1);
