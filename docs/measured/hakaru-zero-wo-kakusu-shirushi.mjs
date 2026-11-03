/* hakaru-zero-wo-kakusu-shirushi.mjs — ★板の「ゼロを 表示しない」の 印★ 2026-09-25
 *
 *  ★★なぜ 要るか（経営者1 の 実測）★★
 *    表示が 実Excel と 違う ★9,163個★ の うち
 *    ★6,743個（74%）は 「Excel は 空・うちは `0`」★ でした。
 *    ⇒`$xl.ActiveWindow.DisplayZeros` を 板ごとに 読んだら
 *      ★15枚中 13枚が False（ゼロを 隠して いる）★
 *    ⇒★計算の 間違いでは ありません＝板の 設定を 読んで いません★
 *
 *  ★★この 道具が する 事★★
 *    ★言われた 事を 信じずに 実物で 確かめます★
 *    ・`.xlsx` ･･･ `xl/worksheets/sheetN.xml` の `<sheetView ･･･ showZeros="0">` を 字で 読む
 *    ・`.xlsb` ･･･ 板の 記録を 舐めて ★板ごとに 同じ 番号の 記録の 生バイトを 並べる★
 *                ⇒★13枚と 2枚で 違う ビット★ が 在れば それが 印
 *    ⇒★番号表を 記憶で 書きません★（記憶「役所の 様式の 表を 記憶で 書くな」）
 *    ⇒★「どのビットか」は ★差で★ 決めます★
 *
 *  ★読むだけ★＝1バイトも 書きません。★マスの 値は 1つも 出しません★。
 *
 *  走らせ方: node docs/measured/hakaru-zero-wo-kakusu-shirushi.mjs <本> [--番号 137]
 */
import path from 'node:path'; import fs from 'node:fs'; import zlib from 'node:zlib';
import crypto from 'node:crypto';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, '../package.json')));
const XlsbEdit = require_(path.join(ROOT, '../lib/xlsb-edit.js'));
const 材料 = process.argv[2] || path.join(ROOT, '../tests/fixtures/cross-sheet-sample.xlsb');

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
    出[名] = 方 === 8 ? zlib.inflateRawSync(体) : 体;
    p = 頭 + 圧;
  }
  return 出;
}

/** ★BIFF12 の 記録★（番号＝可変長・長さ＝可変長） */
function 記録たち(bin, 打ち切り) {
  const 出 = [];
  let p = 0;
  while (p < bin.length) {
    let 番 = bin[p++];
    if (番 & 0x80) 番 = (番 & 0x7f) | ((bin[p++] & 0x7f) << 7);
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
    /* ★頭の 方だけ 見れば 足ります★＝板の 設定は 前の 方に 在ります
         ★切った 事は 字で 言います★（2026-09-25 経営者1 が 24バイトで 切って 嘘を 出した） */
    if (打ち切り && 出.length >= 打ち切り) { 出.打ち切った = true; break; }
  }
  return 出;
}

if (!fs.existsSync(材料)) { console.log('★材料が 有りません★ ' + 材料); process.exit(1); }
const 中 = fs.readFileSync(材料);
console.log('★板の「ゼロを 表示しない」の 印★');
console.log('  材料 ... ' + path.basename(材料) + '（' + 中.length.toLocaleString() + ' バイト）');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex').slice(0, 24));
const 包 = ほどく(中);

/* ══ `.xlsx` の 道＝字で 読める ══ */
const xml = Object.keys(包).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k));
if (xml.length) {
  console.log('  ★`.xlsx` の 道★（' + xml.length + '枚）');
  for (const k of xml.sort()) {
    const s = 包[k].toString('utf8');
    const i = s.indexOf('<sheetView');
    const 頭 = i < 0 ? '(sheetView が 無い)' : s.slice(i, s.indexOf('>', i) + 1);
    console.log('    ' + k.padEnd(28) + ' ' + 頭.slice(0, 150));
  }
}

/* ══ `.xlsb` の 道＝差で 決める ══ */
const bin = Object.keys(包).filter((k) => /^xl\/worksheets\/sheet\d+\.bin$/.test(k));
if (bin.length) {
  const 名指し = [];
  for (let i = 0; i < process.argv.length; i++) {
    if (process.argv[i] === '--番号' && process.argv[i + 1]) 名指し.push(Number(process.argv[i + 1]));
  }
  /* ★板の 名前を 出す★（`workbook.bin` ＋ rels）＝★どの 板が どれか 分かるように★ */
  let 名たち = [];
  try {
    const rels = 包['xl/_rels/workbook.bin.rels'] || 包['xl/_rels/workbook.xml.rels'];
    名たち = XlsbEdit.板たち(包['xl/workbook.bin'], rels ? rels.toString('utf8') : '');
  } catch (e) { 名たち = []; }

  console.log('  ★`.xlsb` の 道★（' + bin.length + '枚）');
  console.log('    ★板ごとに 頭の 40記録を 舐めて 番号ごとの 生バイトを 並べます★');
  const 表 = {};   /* 番号 → [板ごとの 16進] */
  const 板名 = [];
  for (const k of bin.sort((a, b) => (+a.replace(/\D/g, '')) - (+b.replace(/\D/g, '')))) {
    const b = 包[k];
    const 記 = 記録たち(b, 40);
    const 見 = {};
    for (const r of 記) {
      if (見[r.番] !== undefined) continue;                 /* 先に 出た 1本だけ */
      if (r.長 > 32) continue;                              /* 長い 物は 設定では ない */
      見[r.番] = b.slice(r.頭, r.頭 + r.長).toString('hex');
    }
    const な = (名たち.find && 名たち.find((x) => x && x.部品 === k)) || null;
    板名.push((な && な.名) ? な.名 : k.replace('xl/worksheets/', ''));
    for (const n of Object.keys(見)) {
      if (!表[n]) 表[n] = [];
      表[n][板名.length - 1] = 見[n];
    }
  }
  console.log('    板 ... ' + 板名.map((x, i) => (i + 1) + ':' + x).join(' '));
  console.log('');
  console.log('    ★板ごとに 中身が 違う 記録だけ 出します★（★同じ 物は 出しません★）');
  let 出た = 0;
  for (const n of Object.keys(表).map(Number).sort((a, b) => a - b)) {
    const 並 = 表[n];
    const 種 = new Set(並.map((x) => String(x)));
    if (種.size <= 1) continue;
    if (名指し.length && 名指し.indexOf(n) < 0) continue;
    console.log('      ★番号 ' + n + '★');
    for (let i = 0; i < 板名.length; i++) {
      console.log('        板' + String(i + 1).padStart(2) + ' ' + String(板名[i]).slice(0, 10).padEnd(11)
        + ((並[i] || '(無し)').match(/../g) || []).join(' '));
    }
    出た++;
  }
  if (!出た) console.log('      ★どの 記録も 板ごとに 同じでした★');
}

console.log('');
console.log('★この 紙の 見方★');
console.log('  ・経営者1 の 実測 ... ★15枚中 13枚が 「ゼロを 隠す」★（板13・14 だけ 出す）');
console.log('  ⇒★13枚と 2枚で 違う ビットが 在れば それが 印★');
console.log('  ⇒★番号表を 記憶で 書いて いません★＝★差で 決めます★');
