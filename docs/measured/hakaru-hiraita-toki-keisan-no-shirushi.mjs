/* hakaru-hiraita-toki-keisan-no-shirushi.mjs — ★「開いたら 全部 計算しろ」の 印が 在るか★ 2026-09-25
 *
 *  ★★なぜ 要るか（経営者1 の 頼み）★★
 *    「開いた 直後の 計算は 要るのか」を 決めるのに ★場合が 5つ★ 在ります。
 *      ㋐開いた 日で 答えが 変わる 関数 ... ★実物 0個★（経営者1 が 数えた）
 *      ㋑式は 在るのに 答えが 保存されて いない マス ... ★実物 0個★（同上）
 *      ㋒Excel 以外の 道具が 作った 本
 *      ㋓★「開いたら 全部 計算しろ」の 印★ ... ★未測定★  ←★ここ★
 *      ㋔お客さんが 打った 後（開いた 直後では ない）
 *    ⇒★㋓だけが 空いて いました★。`.bin` は 向こうの 道具では 読めません。
 *
 *  ★★何を 見るか★★
 *    `.xlsx` ･･･ `xl/workbook.xml` の `<calcPr .../>`（★字で 読めます★）
 *    `.xlsb` ･･･ `xl/workbook.bin` の 記録を 1つずつ 舐めて ★番号と 長さ★を 出す
 *              ＋ `calcPr` に あたる 記録の ★生の バイト★を 16進で 出す
 *
 *  ★★★正直に 書きます ── 私は 番号を 知りません★★★
 *    MS-XLSB の 番号表を ★記憶で 書きません★（記憶「役所の 様式の 表を 記憶で 書くな」）。
 *    ⇒★この 道具は 「答え」を 出しません★。★材料を 並べるだけ★です。
 *    ⇒★2つの 本を 並べて 「どの 記録が 在る／無い」を 見る★のが 使い道です。
 *    ⇒★ビットの 意味は ここでは 決めません★＝経営者1 が 実Excel で 印を 付けた 本を
 *      作って くれれば ★差が 出た 記録が それ★ と 言えます。
 *
 *  ★読むだけ★＝1バイトも 書きません。★中身（マスの 値）は 出しません★。
 *
 *  走らせ方: node docs/measured/hakaru-hiraita-toki-keisan-no-shirushi.mjs <本> [<本2> ...]
 */
import path from 'node:path'; import fs from 'node:fs'; import zlib from 'node:zlib';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 本たち = process.argv.slice(2);
if (!本たち.length) {
  本たち.push(path.join(ROOT, '../tests/fixtures/cross-sheet-sample.xlsb'));
  本たち.push(path.join(ROOT, '../tests/fixtures/kazari-hiraku3.xlsx'));
}

/** ★包みを ほどく★（借り物に 渡さず 局所ヘッダを 順に 舐める） */
function ほどく(buf) {
  const 出 = {};
  let p = 0;
  while (p + 30 <= buf.length) {
    if (buf.readUInt32LE(p) !== 0x04034b50) break;
    const 方 = buf.readUInt16LE(p + 8);
    const 圧 = buf.readUInt32LE(p + 18);
    const n = buf.readUInt16LE(p + 26);
    const x = buf.readUInt16LE(p + 28);
    const 名 = buf.toString('utf8', p + 30, p + 30 + n);
    const 頭 = p + 30 + n + x;
    const 体 = buf.slice(頭, 頭 + 圧);
    出[名] = 方 === 8 ? zlib.inflateRawSync(体) :体;
    p = 頭 + 圧;
  }
  return 出;
}

/** ★BIFF12 の 記録を 1つずつ★（番号＝可変長・長さ＝可変長） */
function 記録たち(bin) {
  const 出 = [];
  let p = 0;
  while (p < bin.length) {
    let 番 = bin[p++];
    if (番 & 0x80) { 番 = (番 & 0x7f) | ((bin[p++] & 0x7f) << 7); }
    let 長 = 0, 位 = 0, b;
    do {
      if (p >= bin.length) return 出;
      b = bin[p++];
      長 |= (b & 0x7f) << 位;
      位 += 7;
    } while (b & 0x80);
    if (p + 長 > bin.length) break;
    出.push({ 番, 長, 頭: p });
    p += 長;
  }
  return 出;
}

for (const 道 of 本たち) {
  if (!fs.existsSync(道)) { console.log('★材料が 有りません★ ' + 道); continue; }
  const 中 = fs.readFileSync(道);
  console.log('');
  console.log('════ ' + path.basename(道) + '（' + 中.length.toLocaleString() + ' バイト）');
  console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex').slice(0, 24));
  const 包 = ほどく(中);
  const 名たち = Object.keys(包);
  console.log('  包みの 中 ' + 名たち.length + '本');

  if (包['xl/workbook.xml']) {
    const x = 包['xl/workbook.xml'].toString('utf8');
    const i = x.indexOf('<calcPr');
    console.log('  ★`.xlsx` の 道★ `xl/workbook.xml`');
    console.log('    `<calcPr>` ... ' + (i < 0 ? '★在りません★'
      : ('★' + x.slice(i, x.indexOf('>', i) + 1) + '★')));
    console.log('    ⇒`fullCalcOnLoad="1"` が 在れば ★開いたら 全部 計算しろ★ の 印');
  }

  if (包['xl/workbook.bin']) {
    const bin = 包['xl/workbook.bin'];
    const 記 = 記録たち(bin);
    console.log('  ★`.xlsb` の 道★ `xl/workbook.bin`（' + bin.length.toLocaleString() + ' バイト）');
    console.log('    記録 ' + 記.length + '本');
    /* ★番号ごとに 数える★＝2つの 本を 並べた 時に 差が 見えるように */
    const 数 = {};
    for (const r of 記) 数[r.番] = (数[r.番] || 0) + 1;
    const 並 = Object.keys(数).map(Number).sort((a, b) => a - b);
    console.log('    番号（番号x個数）... ' + 並.map((n) => n + 'x' + 数[n]).join(' '));
    /* ★1回しか 出ない 短い 記録★＝決めごとが 入って いる 見込みが 高い
         ⇒★生の バイトを 出します★（マスの 値では ありません＝決めごとです） */
    /* ★★番号を 名指しで 出せます★★＝`--番号 157 --番号 153` の ように */
    const 名指し = [];
    for (let i = 0; i < process.argv.length; i++) {
      if (process.argv[i] === '--番号' && process.argv[i + 1]) 名指し.push(Number(process.argv[i + 1]));
    }
    for (const n of 名指し) {
      const あ = 記.filter((r) => r.番 === n);
      console.log('    ★名指し 番号 ' + n + '★ ... ' + あ.length + '本');
      for (const r of あ.slice(0, 3)) {
        const b2 = bin.slice(r.頭, r.頭 + Math.min(r.長, 64));
        console.log('      長さ ' + r.長 + ' ／ ' + (b2.toString('hex').match(/../g) || []).join(' ')
          + (r.長 > 64 ? ' ･･･' : ''));
      }
    }
    console.log('    ★1回だけ 出る 短い 記録（16バイト以下）の 生バイト★');
    let 出た = 0;
    for (const r of 記) {
      if (数[r.番] !== 1 || r.長 > 16) continue;
      const b = bin.slice(r.頭, r.頭 + r.長);
      console.log('      番号 ' + String(r.番).padStart(4) + ' ／ 長さ ' + String(r.長).padStart(2)
        + ' ／ ' + (b.toString('hex').match(/../g) || []).join(' '));
      出た++;
      if (出た >= 24) { console.log('      ･･･（24本まで）'); break; }
    }
  }
}

console.log('');
console.log('★★この 紙が 言えない 事★★');
console.log('  ・★どの 番号が `calcPr` かは ここでは 決めません★（★番号表を 記憶で 書きません★）');
console.log('  ・★ビットの 意味も 決めません★');
console.log('  ⇒★実Excel で 「開いたら 全部 計算」を 付けた 本と 付けない 本を 並べれば★');
console.log('    ★差が 出た 記録と ビットが それ★ と 言えます（★経営者1 の 持ち場★）');
