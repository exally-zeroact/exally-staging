/* kuraberu-tsutsumi-no-buhin.mjs
 *   -- ★包みの 部品を 前後で 突き合わせる★（95）（2026-09-21）
 *
 *  ★★なぜ★★
 *    Exally1 が 「足した 板も 書き出す」を 書きます。
 *    ⇒私（監査役）が 数えるのは ★「増えた／減った／中身が 変わった」★です。
 *    ⇒★これは Excel を 1回も 立てずに 分かります★（速い・お金 0）
 *    ⇒★実Excel に 開かせるのは その 後★（`Open` が 投げるか＝93）
 *
 *  ★★この 道具が する 事／しない 事★★
 *    する  ... 2つの 包みの ★部品の 名前と 中身（sha256）★を 並べる
 *    しない ... ★中身が 「正しい」かの 判断★（それは 実Excel に 開かせて 見る）
 *
 *  ★★「変わって いない」の 決め方★★
 *    ★ほどいた 後の バイトで 比べます★（★詰め方は 見ません★）
 *      ＝zip の 詰め直しで バイトが 変わっても ★中身は 同じ★事が 在る
 *      ＝`lib/vba.js` の 「中身は同じ・詰め方が違う」と ★同じ 考え方★
 *    ⇒★詰め方の 違いも 別に 数えて 出します★（黙って 同じに しない）
 *
 *  使い方:
 *    node docs/measured/kuraberu-tsutsumi-no-buhin.mjs <前の包み> <後の包み>
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const require_ = createRequire(path.join(ここ, '..', '..', 'package.json'));
const zlib = require_('node:zlib');
const crypto = require_('node:crypto');

/** ★zip を ほどいて 名前 ⇒ {中身, 詰め方} に する★（借り物を 増やしません） */
function 包みを開く(buf) {
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return null;
  const 本数 = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  const 出 = new Map();
  for (let k = 0; k < 本数; k += 1) {
    if (buf.readUInt32LE(p) !== 0x02014b50) return null;
    const 詰め方 = buf.readUInt16LE(p + 10);
    const 圧んだ = buf.readUInt32LE(p + 20);
    const 名長 = buf.readUInt16LE(p + 28);
    const 追長 = buf.readUInt16LE(p + 30);
    const 注長 = buf.readUInt16LE(p + 32);
    const 頭位置 = buf.readUInt32LE(p + 42);
    const 名 = buf.slice(p + 46, p + 46 + 名長).toString('utf8');
    const 名長2 = buf.readUInt16LE(頭位置 + 26);
    const 追長2 = buf.readUInt16LE(頭位置 + 28);
    const 生 = buf.slice(頭位置 + 30 + 名長2 + 追長2, 頭位置 + 30 + 名長2 + 追長2 + 圧んだ);
    const 中身 = 詰め方 === 0 ? 生 : zlib.inflateRawSync(生);
    出.set(名, { 中身, 詰め方, 圧んだ: 圧んだ });
    p += 46 + 名長 + 追長 + 注長;
  }
  return 出;
}

const 印 = (b) => crypto.createHash('sha256').update(b).digest('hex').slice(0, 16);

const 前道 = process.argv[2];
const 後道 = process.argv[3];
if (!前道 || !後道) {
  console.log('使い方: node kuraberu-tsutsumi-no-buhin.mjs <前の包み> <後の包み>');
  process.exit(2);
}
for (const d of [前道, 後道]) {
  if (!fs.existsSync(d)) { console.log('★在りません★ ... ' + d); process.exit(3); }
}
/* ══ ★★材料の 印を 先に 突き合わせます★★ ══（2026-09-22）
     ★`%TEMP%` の 名は 上書きされます★。実際に `exally-tashita-ita.xlsx` の 中身が
     ★別の 材料の 物に 入れ替わって いました★（気づかずに 測って いました）。
     ⇒★違う 材料どうしを 並べると ★相手の 直しの せいに 見える 嘘の 赤★が 出ます★
     ⇒`--前の印` / `--後の印` に sha256 を 渡すと ★合わない なら 走りません★
     ⇒★渡されて いない 時は 「見て いません」と 出します★（★黙って 通さない★） */
const 印を取る = (b) => require_('node:crypto').createHash('sha256').update(b).digest('hex');
const 印引数 = (な) => {
  const i = process.argv.indexOf(な);
  return i > 0 ? String(process.argv[i + 1] || '').toLowerCase() : '';
};
{
  const ま1 = 印引数('--前の印'), ま2 = 印引数('--後の印');
  if (ま1 || ま2) {
    const h1 = 印を取る(fs.readFileSync(前道)), h2 = 印を取る(fs.readFileSync(後道));
    if (ま1 && h1 !== ま1) { console.log('★★前の 材料が 違います★★ 待ち ' + ま1 + ' ／ 実物 ' + h1); process.exit(2); }
    if (ま2 && h2 !== ま2) { console.log('★★後の 材料が 違います★★ 待ち ' + ま2 + ' ／ 実物 ' + h2); process.exit(2); }
    console.log('★印は 渡された 物と 合って います★');
  } else {
    console.log('★★印を 渡されて いません＝すり替わりを 見て いません★★');
  }
}

const 前 = 包みを開く(fs.readFileSync(前道));
const 後 = 包みを開く(fs.readFileSync(後道));
if (!前 || !後) { console.log('★包みが ほどけません★'); process.exit(4); }

console.log('★前★ ' + 前道 + '（部品 ' + 前.size + '本）');
console.log('★後★ ' + 後道 + '（部品 ' + 後.size + '本）');
console.log('');

const 名たち = [...new Set([...前.keys(), ...後.keys()])].sort();
const 増 = [], 減 = [], 変 = [], 同 = [], 詰め違い = [];
for (const な of 名たち) {
  const a = 前.get(な), b = 後.get(な);
  if (!a) { 増.push(な); continue; }
  if (!b) { 減.push(な); continue; }
  if (印(a.中身) !== 印(b.中身)) { 変.push({ な, 前: a.中身.length, 後: b.中身.length }); continue; }
  同.push(な);
  if (a.圧んだ !== b.圧んだ) 詰め違い.push(な);
}

const 出す = (札, 並び, 書き) => {
  console.log('★' + 札 + '★ ' + 並び.length + '本');
  for (const x of 並び.slice(0, 40)) console.log('    ' + 書き(x));
  if (並び.length > 40) console.log('    ...（あと ' + (並び.length - 40) + '本）');
};
出す('増えた 部品', 増, (x) => x + '  ' + 後.get(x).中身.length + 'B');
/* ══ ★★「わざと 消した 物」の 扱い★★ ══（2026-09-22・Exally1 の 問い）
     ★彼の 言い分★「消して 良い 物を 道具に 教えると ★本当に 消えた 時も 通る★」
     ⇒★その通りです★。だから ★道具に 覚えさせません★。
     ★でも 「人が 毎回 見る」も 門では ありません★（人は 見落とします）
     ⇒★★その 場で 名指しして 宣言する★★ 形に しました
         `--消える予定 xl/worksheets/binaryIndex1.bin`
       ・★宣言した 物★ ... 「★わざと 消した★」と 出す（赤に しない）
       ・★宣言して いない 物★ ... ★今まで 通り 赤★
       ・★宣言したのに 消えて いない★ ... ★これも 赤★（★宣言が 古い★のを 見つける）
     ⇒★覚えさせない／黙って 通さない／宣言は 命令の 字に 残る★ */
const 消える予定 = (() => {
  const 出 = [];
  for (let i = 0; i < process.argv.length; i += 1) {
    if (process.argv[i] === '--消える予定') 出.push(String(process.argv[i + 1] || ''));
  }
  return 出.filter(Boolean);
})();
const わざと = 減.filter((x) => 消える予定.indexOf(x) >= 0);
const 本当に減った = 減.filter((x) => 消える予定.indexOf(x) < 0);
const 宣言が古い = 消える予定.filter((x) => 減.indexOf(x) < 0);

出す('★わざと 消した 部品（宣言あり）★', わざと, (x) => x + '  ' + 前.get(x).中身.length + 'B');
出す('★★宣言して いないのに 減った 部品★★', 本当に減った, (x) => x + '  ' + 前.get(x).中身.length + 'B');
出す('★★消える予定なのに 残って いる（宣言が 古い）★★', 宣言が古い, (x) => x);
出す('★中身が 変わった 部品★', 変, (x) => x.な + '  ' + x.前 + 'B ⇒ ' + x.後 + 'B');
console.log('★中身が 同じ 部品★ ' + 同.length + '本'
  + (詰め違い.length ? '（うち ★詰め方だけ 違う★ ' + 詰め違い.length + '本）' : ''));
/* ══ ★★`.xlsb` の 「字の マス」を 数えます★★ ══（2026-09-22）
     ★`.xlsb` に 足した 板は 今 ★数しか 書けません★★（`sharedStrings.bin` を 触れない）
     ⇒★字を 打った つもりの マスは 黙って 落ちます★
     ⇒★落ちるのは 狙い通りですが 「落ちた」事が 見えないと 嘘に なります★
     ⇒★増えた `.bin` の 板に 字の マス（記録 7）が 何本 在るかを 出します★
       ＝0本なら ★字は 1つも 入って いません★と はっきり 言う
     ★記録の 番号★ ... 2 ＝数の マス ／ 7 ＝字の マス（★実測・85 の 紙★） */
function 記録に割る(b) {
  const 出 = [];
  let p = 0;
  const 継ぐ = (q, 最大) => {
    let v = 0;
    for (let i = 0; i < 最大; i += 1) {
      const c = b[q + i];
      v |= (c & 0x7f) << (7 * i);
      if ((c & 0x80) === 0) return [v >>> 0, i + 1];
    }
    return null;
  };
  while (p < b.length) {
    const 番 = 継ぐ(p, 2); if (!番) break;
    const 長 = 継ぐ(p + 番[1], 4); if (!長) break;
    const h = p + 番[1] + 長[1];
    if (h + 長[0] > b.length) break;
    出.push(番[0]);
    p = h + 長[0];
  }
  return 出;
}
const 足した板 = 増.filter((x) => /^xl\/worksheets\/sheet\d+\.bin$/.test(x));
if (足した板.length) {
  console.log('');
  console.log('★★足した 板の 中身（`.xlsb`）★★');
  for (const な of 足した板) {
    const 並 = 記録に割る(後.get(な).中身);
    const 数 = 並.filter((n) => n === 2).length;
    const 字 = 並.filter((n) => n === 7).length;
    console.log('  ' + な + ' ... 記録 ' + 並.length + '本 ／ ★数の マス ' + 数
      + '本★ ／ ★字の マス ' + 字 + '本★'
      + (字 === 0 ? '（★字は 1つも 入って いません★）' : ''));
  }
}

console.log('');
console.log('★★この 道具は 「開けるか」を 見て いません★★');
console.log('  ⇒実Excel が 投げないかは `toru-jitsu-excel-ga-shuufuku-shita-ka.ps1`（93）で 見ます');
if (本当に減った.length) {
  console.log('');
  console.log('★★宣言して いない 部品が 減りました★★');
  console.log('  ⇒わざとなら `--消える予定 <名>` と ★命令の 字に 書いて ください★');
  process.exit(1);
}
if (宣言が古い.length) {
  console.log('');
  console.log('★★消える予定と 言われた 物が 残って います＝宣言が 古いです★★');
  process.exit(1);
}
