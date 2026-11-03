/* hakaru-haba-no-de-dokoro.mjs — ★列の 幅は どの 数から 来て いるか★ 2026-09-24
 *
 *  ★★司さんの 言葉★★
 *    「★読み込んだ ファイル 見切れとん とかも 自動で 調整しろ★」
 *    「★は？ Excel内で 見切れてない とこが 見切れとるけん いよんやろが★」
 *    ⇒★Excel と 違う 事を しろ では ない★＝★Excel に 合って いない★＝不具合。
 *
 *  ★★これは 「直す」道具では なく 「数える」道具★★
 *    ★見立てで 直さない★＝まず ★借り物が 何を くれて いるか★ を 出す。
 *      `ws['!cols'][i]` の `width` / `wch` / `wpx` / `customWidth`
 *    そして ★今の 計算★ と ★足りて いない かも しれない 5点★ を 並べる。
 *
 *  ★★今の 計算（`js/book-open.js:516 幅を点に()`）★★
 *      `width` が 在る ⇒ ★round(width × 字幅)★        （★余白 5 を 足さない★）
 *      無ければ `wch`  ⇒ ★round(wch × 字幅 + 5)★      （★こちらは 足す★）
 *      どちらも 無い   ⇒ 0 ⇒ ★標準の点 = round(8.43 × 字幅 + 5)★（★足す★）
 *    ⇒★同じ 8.43 でも 道が 違うと 5点 ずれます★
 *
 *  ★読むだけ★＝材料を 1バイトも 書きません。
 *
 *  走らせ方:
 *    node docs/measured/hakaru-haba-no-de-dokoro.mjs [材料の道]
 */
import path from 'node:path';
import fs from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, '../package.json')));
const XLSX = require_(path.join(ROOT, '../lib/xlsx.full.min.js'));

const 材料 = process.argv[2] || path.join(ROOT, '../tests/fixtures/cross-sheet-sample.xlsb');
if (!fs.existsSync(材料)) { console.log('★材料が 有りません★ ' + 材料); process.exit(1); }

const 字幅 = 8;          /* ★游ゴシック 11pt の 実測（webkit 8.157 ⇒ 床 8）★ */
const 標準の字数 = 8.43;

console.log('★列の 幅は どの 数から 来て いるか★');
console.log('  材料 ... ' + path.basename(材料) + '（' + fs.statSync(材料).size.toLocaleString() + ' バイト）');
console.log('  ★字幅 ' + 字幅 + '点★（游ゴシック 11pt・webkit 実測 8.157 の 床）');
console.log('  ★標準の点 = round(8.43 × ' + 字幅 + ' + 5) = ' + Math.round(標準の字数 * 字幅 + 5) + '★');
console.log('');

const wb = XLSX.read(new Uint8Array(fs.readFileSync(材料)), { type: 'array', cellStyles: true });

let 板数 = 0, 列数 = 0;
const 数え = { width在り: 0, wchだけ: 0, wpx在り: 0, 何も無い: 0, custom: 0 };
const 見本 = [];

for (const 名 of wb.SheetNames) {
  const ws = wb.Sheets[名];
  if (!ws) continue;
  板数++;
  const cols = ws['!cols'] || [];
  for (let i = 0; i < cols.length; i++) {
    const c = cols[i];
    if (!c) { 数え.何も無い++; continue; }
    列数++;
    if (typeof c.wpx === 'number') 数え.wpx在り++;
    if (c.customWidth) 数え.custom++;
    if (typeof c.width === 'number' && c.width > 0) 数え.width在り++;
    else if (typeof c.wch === 'number' && c.wch > 0) 数え.wchだけ++;
    else 数え.何も無い++;
    if (見本.length < 12) {
      const 今 = (typeof c.width === 'number' && c.width > 0) ? Math.round(c.width * 字幅)
        : (typeof c.wch === 'number' && c.wch > 0) ? Math.round(c.wch * 字幅 + 5) : 0;
      const 足した = (typeof c.width === 'number' && c.width > 0) ? Math.round(c.width * 字幅 + 5) : 今;
      見本.push({
        板: 名, 列: i,
        width: c.width, wch: c.wch, wpx: c.wpx, custom: !!c.customWidth,
        今: 今, 五を足すと: 足した, 差: 足した - 今,
      });
    }
  }
}

console.log('★数えた 所★ 板 ' + 板数 + '枚 ／ 幅を 持つ 列 ' + 列数 + '本');
console.log('  width が 在る ... ' + 数え.width在り + '本（★余白 5 を 足して いない 道★）');
console.log('  wch だけ ....... ' + 数え.wchだけ + '本（余白 5 を 足す 道）');
console.log('  wpx が 在る ..... ' + 数え.wpx在り + '本（★借り物が 点で くれて いる★）');
console.log('  何も 無い ....... ' + 数え.何も無い + '本（標準の点）');
console.log('  customWidth ..... ' + 数え.custom + '本');
console.log('');
console.log('★見本（先頭 ' + 見本.length + '本）★');
console.log('  板              列  width     wch      wpx   今の点  +5した点  差');
for (const x of 見本) {
  console.log('  ' + String(x.板).slice(0, 14).padEnd(15)
    + String(x.列).padStart(3)
    + String(x.width === undefined ? '-' : x.width).padStart(9)
    + String(x.wch === undefined ? '-' : Math.round(x.wch * 100) / 100).padStart(9)
    + String(x.wpx === undefined ? '-' : x.wpx).padStart(8)
    + String(x.今).padStart(8) + String(x.五を足すと).padStart(10)
    + String(x.差).padStart(4));
}
console.log('');
console.log('★★この 紙が 言えない 事★★');
console.log('  ・★実Excel が 何点で 描いて いるかは ここでは 分かりません★（COM は 私の 持ち場では ない）');
console.log('  ・★見切れの 因が これだとは まだ 言えません★＝★数を 並べただけ★');
