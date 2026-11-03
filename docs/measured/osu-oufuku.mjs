/* osu-oufuku.mjs — ★うちで 書き出す／実Excel の 結果と 突き合わせる★（2026-09-10）
 *
 *  ★★なぜ 要るか（★今日まで 一度も やって いませんでした★）★★
 *    今まで … ★実Excel に 打たせて 答えを 合わせる★（Excel → うち）
 *    やって いなかった … ★うちが 書いた 物を Excel で 開き直す★（うち → Excel）
 *    ⇒ 司さん（2026-09-10）「★Exally でも Excel でも 使えるように 確かめながら やってるか？★」
 *
 *  ★★この 台の 数は 画面の 数では ありません★★
 *    ここは ★書き出しの 道（GridXlsx → XlsxIO）★を 押すだけです。
 *    ★画面の 事を 言いたい時は ブラウザで 押す★
 *
 *  使い方:
 *    node docs/measured/osu-oufuku.mjs           … ★うちで 書き出す★（xlsx を 作る）
 *    pwsh -File docs/measured/toru-oufuku.ps1    … ★実Excel で 開いて 測る★
 *    node docs/measured/osu-oufuku.mjs --awase   … ★突き合わせる★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const ここ = path.join(ROOT, 'docs/measured');
const ファイル = path.join(ここ, 'oufuku-2026-09-10.xlsx');
const 紙 = path.join(ここ, 'golden-oufuku-2026-09-10.tsv');
const 出す先 = path.join(ここ, 'golden-oufuku-awase-2026-09-10.tsv');
const 突き合わせ = process.argv.includes('--awase');

const require_ = createRequire(path.join(ROOT, 'package.json'));

/* ══ ★材料（★今日 直した 物 全部★）★ ══
   マス ／ 打つ 字 ／ ★うちの 答え（本番の 道で 出した 物）★ ／ 何を 見て いるか
   ★実Excel が 式を 書き換える 物は ここに 書いて おく★（次に 見た 人が 穴だと 思わない為） */
const 材料 = [
  { マス: 'D1', 式: '=XIRR(A1:A3,B1:B3)', 何: 'お金の 利回り（PR #59）' },
  { マス: 'D2', 式: '=MIRR(A1:A3,0.1,0.12)', 何: 'お金の 利回り（PR #60）' },
  { マス: 'D3', 式: '=XNPV(0.1,A1:A3,B1:B3)', 何: 'お金の 今の 値（PR #59）' },
  { マス: 'D4', 式: '=1.64E-14', 何: '指数の 字（PR #62）',
    式が変わる: '=0.0000000000000164',
    訳: '★実Excel は 指数の 字を 打つと 式そのものを 平らな 数に 直して 持ちます★＝★穴では ない★' },
  { マス: 'D5', 式: '=1/3', 何: '画面に 出る 字（PR #63）' },
  { マス: 'D6', 式: '=206800/1.1', 何: '税抜き（PR #63）' },
  { マス: 'G1', 式: '=SORT(E1:E5)', 何: '溢れの 頭（PR #65）' },
];
/* ★溢れた 先（式は 無く 値だけ 渡る＝それで 正しい）★ */
const 溢れ先 = [
  { マス: 'G2', 値: '2' }, { マス: 'G3', 値: '3' }, { マス: 'G4', 値: '4' }, { マス: 'G5', 値: '5' },
];

function マスを分ける(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  const c = m[1].split('').reduce((s, ch) => s * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
  return { r: Number(m[2]) - 1, c: c };
}

if (!突き合わせ) {
  /* ══ ★うちで 書き出す★ ══ */
  global.XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
  const GX = require_(path.join(ROOT, 'lib/grid-xlsx.js'));
  const IO = require_(path.join(ROOT, 'lib/xlsx-io.js'));

  const 板 = { data: {}, name: 'Sheet1' };
  const 置く = (a, o) => { const p = マスを分ける(a); 板.data[p.r + ',' + p.c] = o; };

  /* ★お金の 流れと 日付★ */
  [[-1000, 45292], [600, 45383], [700, 45474]].forEach(([v, d], i) => {
    置く('A' + (i + 1), { v: v }); 置く('B' + (i + 1), { v: d });
  });
  /* ★並べ替えの 種★ */
  [3, 1, 5, 2, 4].forEach((v, i) => 置く('E' + (i + 1), { v: v }));
  /* ★式★ */
  for (const m of 材料) 置く(m.マス, { v: m.式, f: m.式, d: '' });
  /* ★溢れた 先（画面が 作る マス）★ */
  for (const m of 溢れ先) { const p = マスを分ける(m.マス); 板.data[p.r + ',' + p.c] = { d: m.値, _溢れ元: '0,6' }; }

  const buf = IO.writeBook(GX.gridToBook([板]));
  fs.writeFileSync(ファイル, Buffer.from(buf));
  console.log('★うちで 書き出した … ' + ファイル + '（' + Buffer.from(buf).length + ' バイト）★');
  console.log('');
  console.log('★次に これを 走らせて ください★');
  console.log('  pwsh -NoProfile -File docs/measured/toru-oufuku.ps1');
  console.log('  node docs/measured/osu-oufuku.mjs --awase');
} else {
  /* ══ ★突き合わせる★ ══ */
  if (!fs.existsSync(紙)) {
    console.error('★先に `pwsh -File docs/measured/toru-oufuku.ps1` を 走らせて ください★');
    process.exit(2);
  }
  const 実 = new Map();
  for (const l of fs.readFileSync(紙, 'utf-8').split(/\r?\n/)) {
    if (!l || l.startsWith('#')) continue;
    const c = l.split('\t');
    if (c.length < 5) continue;
    実.set(c[0], { 式: c[2], 字: c[3], 答: c[4] });
  }
  if (!実.size) { console.error('★紙が 読めない★'); process.exit(2); }

  const 行 = [];
  行.push('# ★うちが 書いた 物を 実Excel で 開き直して 突き合わせた★（2026-09-10）');
  行.push('#');
  行.push('# ★見るのは 3つ★ ①式 ②出る字 ③答え（★答えだけ 見ない★）');
  行.push('# ★実Excel が わざと 書き換える 物は 材料に 書いて 在る★（穴では ない）');
  行.push('#');
  行.push(['# マス', '打った 字', '実Excel の 式', '出る字', '判じ', '何を 見て いるか'].join('\t'));

  let 合 = 0, 違 = 0;
  const 実物 = [];
  for (const m of 材料) {
    const e = 実.get(m.マス);
    if (!e) { console.error('★紙に 無い … ' + m.マス + '★'); process.exit(2); }
    const 期待 = m.式が変わる || m.式;
    const 式合う = e.式 === 期待;
    /* ★式が 落ちて いないか★＝これが 一番 大事（値だけ 残る 事故） */
    const 式が在る = e.式 && e.式.charAt(0) === '=';
    let 判;
    if (!式が在る) { 判 = '★式が 落ちた★'; }
    else if (式合う) { 判 = m.式が変わる ? '合った（★Excel が 書き換えて 正しい★）' : '合った'; }
    else { 判 = '★式が 違う★'; }
    if (判.startsWith('合った')) 合++; else { 違++; 実物.push(m.マス + ' 打った=' + m.式 + ' ／ Excel=' + e.式); }
    行.push([m.マス, m.式, e.式, e.字, 判, m.何 + (m.訳 ? '｜' + m.訳 : '')].join('\t'));
  }
  /* ★溢れた 先＝式は 無く 値だけ 渡るのが 正しい★ */
  for (const m of 溢れ先) {
    const e = 実.get(m.マス);
    if (!e) continue;
    const 良い = e.答 === m.値;
    if (良い) 合++; else { 違++; 実物.push(m.マス + ' 値=' + e.答 + '（' + m.値 + ' のはず）'); }
    行.push([m.マス, '(溢れた 先)', e.式, e.字, 良い ? '合った（★値が 渡る★）' : '★値が 違う★',
      '溢れの 先（PR #65）｜★式は 無く 値だけ 渡るのが 正しい★'].join('\t'));
  }

  行.push('#');
  行.push('# ★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
  fs.writeFileSync(出す先, 行.join('\n') + '\n', 'utf-8');

  console.log('');
  console.log('★★締め★★ 全 ' + (合 + 違) + '本 ／ 合った ' + 合 + ' ／ ★違う ' + 違 + '★');
  if (実物.length) {
    console.log('');
    console.log('★合わない 物の 実物★');
    実物.forEach((x) => console.log('  ' + x));
  }
  console.log('');
  console.log('★書いた … ' + 出す先 + '★');
  if (違) process.exit(1);
}
