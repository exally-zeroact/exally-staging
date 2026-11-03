/* jitsubutsu-vba-scan.mjs — ★実物のマクロを 数える（読むだけ）★
 * =============================================================================
 * ★決まり（指示役 2026-08-28）★
 *   ・★読むだけ★＝1バイトも 書かない・repoに 写さない
 *   ・★ファイル名も 中身の字も 出さない★（客の会社名が 入っているため）
 *     出すのは ★本数★と ★なぜ★だけ。
 *   ・★VBAは 動かさない★。読んで 数えるだけ。
 *   ・★読めない物は「未測定」★＝0本と 言わない。
 *
 * 何を出すか
 *   ①マクロが 入っているファイルは 何本か（母数を 先に書く）
 *   ②その中で ★正しく読めた★のは 何本か（＝もう一度 圧縮して 元のバイトと 一致）
 *   ③手続きは 何本で、可否（そのまま できる／やり方を かえる／別の道）の 内訳
 *   ④どの分類が 多いか（★うちの表の言葉だけ★）
 *
 * 使い方: node scripts/jitsubutsu-vba-scan.mjs "C:/Users/zeroa/OneDrive"
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const ZS = require_(path.join(ROOT, 'lib/zip-surgeon.js'));
const V = require_(path.join(ROOT, 'lib/vba.js'));
const M = require_(path.join(ROOT, 'lib/vba-mikata.js'));
const TJ = require_(path.join(ROOT, 'lib/vba-tejun.js'));

const 元 = process.argv[2];
if (!元) { console.log('★どこを見るか 教えてください★'); process.exit(1); }

/* ★読むだけ★＝この道具は 書き込みの関数を 1回も 呼ばない */
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
console.log('[実物のマクロ] ★読むだけ★ ' + ファイル.length + '本');
console.log('  ★出すのは 本数と「なぜ」だけ★（ファイル名・コードの中身は 1文字も 出しません）');
console.log('');

const 数 = {
  本: ファイル.length, 開けない: 0,
  マクロ在り: 0, 読めた: 0, 読めない: 0,
  モジュール: 0, 正しく読めた: 0, 未測定: 0, 詰め方も同じ: 0, 名前一致: 0, 往復: 0,
  手続き: 0, 手順が取れた手続き: 0, 手順: 0, 手順の効く行: 0, 手順の読めない行: 0,
};
const 可否 = { 'できる': 0, 'かえる': 0, 'べつの道': 0, 'わからない': 0 };
const 分類ごと = {};
const なぜ別 = {};
const 字づかい別 = {};

for (const f of ファイル) {
  let bytes;
  try { bytes = new Uint8Array(fs.readFileSync(f)); } catch (e) { 数.開けない++; continue; }
  if (bytes[0] !== 0x50 || bytes[1] !== 0x4B) continue;      /* zip でない（.xls）＝ここでは 見ない */
  let z;
  try { z = ZS.read(bytes); } catch (e) { 数.開けない++; continue; }
  if (!z.has('xl/vbaProject.bin')) continue;
  数.マクロ在り++;
  let bin;
  try { bin = await z.bytes('xl/vbaProject.bin'); }
  catch (e) { 数.読めない++; なぜ別['入れ物を 取り出せない'] = (なぜ別['入れ物を 取り出せない'] || 0) + 1; continue; }

  const r = V.読む(bin, XLSX.CFB);
  if (!r.ok) { 数.読めない++; なぜ別[r.なぜ] = (なぜ別[r.なぜ] || 0) + 1; continue; }
  数.読めた++;
  数.モジュール += r.モジュール.length;
  数.正しく読めた += r.確かめ.一致;
  数.詰め方も同じ += r.確かめ.詰め方も同じ;
  数.名前一致 += r.確かめ.名前一致;
  数.往復 += r.確かめ.往復;
  for (const m of r.モジュール) {
    if (!m.確か) { 数.未測定++; なぜ別[m.なぜ] = (なぜ別[m.なぜ] || 0) + 1; }
    字づかい別[m.字づかい || '（無し）'] = (字づかい別[m.字づかい || '（無し）'] || 0) + 1;
  }

  const 見 = M.見立てる(r.モジュール);
  数.手続き += 見.本数;
  for (const k of Object.keys(見.数)) 可否[k] = (可否[k] || 0) + 見.数[k];
  for (const v of 見.手続き) for (const a of v.分類) 分類ごと[a.名] = (分類ごと[a.名] || 0) + 1;
  /* ★③レシピにして 会社に残す★＝手順が どれだけ 取り出せるか（★取れない事も 数える★） */
  for (const v of 見.手続き) {
    const t = TJ.取り出す(v);
    数.手順の効く行 += t.数.変える行;
    数.手順の読めない行 += t.数.読めない行;
    数.手順 += t.数.手順;
    if (t.数.手順) 数.手順が取れた手続き++;
  }
}

const 率 = (a, b) => (b ? (Math.round(a / b * 1000) / 10) + '%' : '—');

console.log('★母数①★ 見たファイル … ' + 数.本 + '本（開けなかった ' + 数.開けない + '本）');
console.log('★母数②★ マクロが 入っている … ' + 数.マクロ在り + '本（' + 率(数.マクロ在り, 数.本) + '）');
console.log('  そのうち 読めた … ' + 数.読めた + '本 ／ 読めない ' + 数.読めない + '本');
console.log('');
console.log('★母数③★ モジュール … ' + 数.モジュール + '本');
console.log('  ★正しく読めた（往復＋名前＋字づかい の 3つとも）★ … ' + 数.正しく読めた
  + '本（' + 率(数.正しく読めた, 数.モジュール) + '）');
console.log('   ・往復で 同じバイトに 戻る … ' + 数.往復 + '本（' + 率(数.往復, 数.モジュール) + '）');
console.log('   ・目次の名前と 中身の1行目が 合う … ' + 数.名前一致 + '本（' + 率(数.名前一致, 数.モジュール) + '）');
console.log('   ・元のバイトと 1バイトずつ 同じ … ' + 数.詰め方も同じ + '本（' + 率(数.詰め方も同じ, 数.モジュール)
  + '）＝Excelの詰め方と うちの詰め方は 違う（中身が 違うわけではない）');
console.log('  未測定（正しく読めたと 言えない） … ' + 数.未測定 + '本');
console.log('');
console.log('★母数④★ 手続き（Sub / Function） … ' + 数.手続き + '本');
console.log('  そのまま できる … ' + 可否['できる'] + '本（' + 率(可否['できる'], 数.手続き) + '）');
console.log('  やり方を かえる … ' + 可否['かえる'] + '本（' + 率(可否['かえる'], 数.手続き) + '）');
console.log('  別の道が いる  … ' + 可否['べつの道'] + '本（' + 率(可否['べつの道'], 数.手続き) + '）');
console.log('  読み取れない   … ' + 可否['わからない'] + '本（' + 率(可否['わからない'], 数.手続き) + '）');

console.log('');
console.log('★母数⑤★ 手順が 取り出せた手続き … ' + 数.手順が取れた手続き + '本／' + 数.手続き + '本（'
  + 率(数.手順が取れた手続き, 数.手続き) + '）');
console.log('  取り出せた手順 … ' + 数.手順 + 'つ');
console.log('  表を 変える行 … ' + 数.手順の効く行 + '行（そのうち 読み取れない ' + 数.手順の読めない行 + '行）');

const 並び = Object.keys(分類ごと).sort((a, b) => 分類ごと[b] - 分類ごと[a]);
if (並び.length) {
  console.log('');
  console.log('★何をしているか（多い順・のべ）★');
  for (const k of 並び) console.log('  ' + k + ' … ' + 分類ごと[k] + '本');
}
const なぜ並び = Object.keys(なぜ別).sort((a, b) => なぜ別[b] - なぜ別[a]);
if (なぜ並び.length) {
  console.log('');
  console.log('★読めなかった／確かと言えなかった 理由★');
  for (const k of なぜ並び) console.log('  ' + k.replace(/★/g, '') + ' … ' + なぜ別[k] + '件');
}
const 字並び = Object.keys(字づかい別).sort((a, b) => 字づかい別[b] - 字づかい別[a]);
if (字並び.length) {
  console.log('');
  console.log('★字づかい（コードの 字の種類）★');
  for (const k of 字並び) console.log('  ' + k + ' … ' + 字づかい別[k] + '本');
}
console.log('');
