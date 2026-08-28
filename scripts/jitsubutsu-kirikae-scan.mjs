/* jitsubutsu-kirikae-scan.mjs — ★うちの実物で「切り替えが効くか」を 数える★
 * =============================================================================
 * ★指示役の指示（2026-08-28）★
 *   ・★150本 全部で つかえる／つかえない を 数える★（つかえない物は ★なぜ★ も）
 *   ・★読むだけ★＝1バイトも 書かない／repo に コピーしない
 *   ・★中身の数字を 報告に出さない★＝出すのは ★本数と「なぜ」だけ★（客の情報だから）
 *
 * ★この道具が 守る事★
 *   ・開くのは ★読み取りだけ★（fs.readFileSync）。書き込みの口を 1つも 持たない。
 *   ・出すのは ★シートの形の話だけ★（行数・列数・別々の日の数・見出しが 何本 在るか）。
 *     ★見出しの中身（人の名前）も 金額も 1文字も 出さない★。
 *   ・ファイル名も 出さない（★客の会社名が 入っているため★）。出すのは ★入れ物の名前★だけ。
 *
 * 使い方: node scripts/jitsubutsu-kirikae-scan.mjs "C:/Users/zeroa/OneDrive"
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const Kirikae = require_(path.join(ROOT, 'lib/kirikae.js'));

const 元 = process.argv[2];
if (!元) { console.log('★どこを見るか 教えてください★'); process.exit(1); }

/* ★読むだけ★＝この道具は 書き込みの関数を 使わない（fs.writeFileSync は 1回も 呼ばない） */
function 集める(dir, 出) {
  let 中;
  try { 中 = fs.readdirSync(dir, { withFileTypes: true }); } catch (e) { return 出; }
  for (const x of 中) {
    const p = path.join(dir, x.name);
    if (x.isDirectory()) {
      if (/node_modules|\.git|\$RECYCLE/i.test(x.name)) continue;
      集める(p, 出);
    } else if (/\.(xlsx|xlsm|xlsb|xls)$/i.test(x.name) && !/^~\$/.test(x.name)) {
      出.push(p);
    }
  }
  return 出;
}

const ファイル = 集める(元, []).sort();
console.log('');
console.log('[実物の切り替え] ★読むだけ★ ' + ファイル.length + '本');
console.log('  ★出すのは 本数と「なぜ」だけ★（ファイル名・見出しの中身・金額は 1文字も 出しません）');

const 数 = { 本: 0, 開けた: 0, 開けない: 0, シート: 0, 出る: 0, 出ない: 0 };
const なぜ別 = {};
const 入れ物ごと = {};
const 出るシートの形 = [];

for (const f of ファイル) {
  数.本++;
  const 入れ物 = path.dirname(f).replace(元, '').replace(/^[\\/]/, '') || '（いちばん上）';
  入れ物ごと[入れ物] = 入れ物ごと[入れ物] || { 本: 0, 出る本: 0 };
  入れ物ごと[入れ物].本++;
  let wb = null;
  try {
    const bytes = new Uint8Array(fs.readFileSync(f));      /* ★読むだけ★ */
    wb = XLSX.read(bytes, { type: 'array', cellFormula: true });
    数.開けた++;
  } catch (e) {
    数.開けない++;
    なぜ別['開けなかった'] = (なぜ別['開けなかった'] || 0) + 1;
    continue;
  }
  let この本で出る = 0;
  for (const name of wb.SheetNames) {
    数.シート++;
    const ws = wb.Sheets[name] || {};
    const data = {};
    for (const k of Object.keys(ws)) {
      if (k[0] === '!') continue;
      const a = XLSX.utils.decode_cell(k);
      const c = ws[k];
      data[a.r + ',' + a.c] = { v: c.v, d: c.w };
    }
    let 姿 = null;
    try { 姿 = Kirikae.見る({ name: name, data: data }); }
    catch (e) { 姿 = { つかえる: false, なぜ: '読めませんでした' }; }
    if (姿.つかえる) {
      数.出る++; この本で出る++;
      /* ★形の話だけ★＝人の名前も 金額も 出さない */
      出るシートの形.push({ 別々の日: 姿.別々の日, 人の列: 姿.人たち.length, 月: 姿.月たち.length });
    } else {
      数.出ない++;
      const なぜ = String(姿.なぜ || '（理由なし）').replace(/（別々の日が \d+日）/, '（別々の日が 少ない）');
      なぜ別[なぜ] = (なぜ別[なぜ] || 0) + 1;
    }
  }
  if (この本で出る) 入れ物ごと[入れ物].出る本++;
}

console.log('');
console.log('── 実測（★母数を 先に書く★） ──');
console.log('  ファイル ' + 数.本 + '本（開けた ' + 数.開けた + '／開けなかった ' + 数.開けない + '）');
console.log('  ★母数① 全シート ' + 数.シート + '枚★ … 効く ' + 数.出る + '枚／効かない ' + 数.出ない + '枚');
/* ★本当の打率＝「切り替えが要る表」だけを 母数にする★（2026-08-28 指示役）
   ＝請求書・見積は そもそも 切り替えが 要らない紙。数に入れると 打率が 嘘になる。 */
const 日付の列が無い = なぜ別['日付の列が ありません'] || 0;
const 小さすぎ = なぜ別['表が 小さすぎます'] || 0;
const 表の数 = 数.シート - 日付の列が無い - 小さすぎ;
console.log('  ★母数② 日付の列が在る表 ' + 表の数 + '枚★'
  + '（全 ' + 数.シート + ' − 日付の列が無い ' + 日付の列が無い + ' − 小さすぎ ' + 小さすぎ + '）');
console.log('  　… ★効く ' + 数.出る + '枚（' + Math.round(数.出る / Math.max(1, 表の数) * 1000) / 10 + '%）★'
  + '／効かない ' + (表の数 - 数.出る) + '枚（日ごとの表でない）');
const 効く本 = Object.values(入れ物ごと).reduce((a, x) => a + x.出る本, 0);
console.log('  ★1枚でも 効くファイル … ' + 効く本 + '本 / ' + 数.本 + '本★');
console.log('');
console.log('── 入れ物ごと（★名前は 入れ物まで★） ──');
for (const [k, v] of Object.entries(入れ物ごと).sort((a, b) => b[1].本 - a[1].本)) {
  console.log('  ' + k + ' … ' + v.本 + '本のうち ★' + v.出る本 + '本★で 効く');
}
console.log('');
console.log('── 効かない理由（多い順） ──');
for (const [k, v] of Object.entries(なぜ別).sort((a, b) => b[1] - a[1])) {
  console.log('  ' + v + '枚 … ' + k);
}
if (出るシートの形.length) {
  const 中 = (a) => a.slice().sort((x, y) => x - y)[Math.floor(a.length / 2)];
  console.log('');
  console.log('── 効くシートの形（★中身は出さない・形だけ★） ──');
  console.log('  別々の日 … 真ん中 ' + 中(出るシートの形.map((x) => x.別々の日)) + '日'
    + '（いちばん少ない ' + Math.min(...出るシートの形.map((x) => x.別々の日))
    + '／多い ' + Math.max(...出るシートの形.map((x) => x.別々の日)) + '）');
  console.log('  人の列 … 真ん中 ' + 中(出るシートの形.map((x) => x.人の列)) + '本');
  console.log('  月 … 真ん中 ' + 中(出るシートの形.map((x) => x.月)) + 'か月');
}
console.log('');
