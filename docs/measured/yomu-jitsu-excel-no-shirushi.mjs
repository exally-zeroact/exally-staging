/* yomu-jitsu-excel-no-shirushi.mjs
 *   -- ★実Excel 自身が 溢れる 式に 付ける 印を 生の 字から 読む★（51・52）（2026-09-20）
 *
 *  ★★なぜ★★
 *    うちが `_xlfn.` を 付ける／付けないの 物差しは ★実Excel★です（記憶）。
 *    ⇒★実Excel に 打たせて 保存させた 物を 読み、印を 1つずつ 写します★
 *
 *  ★★読む 物（★2本だけ★／実Excel が 作った 物）★★
 *    `%TEMP%\exally-jitsu-excel-kobore.xlsx`  ＝ 51（SEQUENCE/SORT/UNIQUE/FILTER/TRANSPOSE/MAKEARRAY/SUM）
 *    `%TEMP%\exally-jitsu-excel-meibo11.xlsx` ＝ 52（うちの 名簿に 無い 11個）
 *    ・★どちらも `tsukuru-jitsu-excel-*.ps1` が 作った 物★
 *    ・★司さんの 実物の 名は 1文字も 在りません★
 *    ・★読むだけ★（1バイトも 書きません）
 *
 *  ★★書かない 事★★
 *    ★「うちが 正しいか」の 判じは ここでは しません★
 *    ＝★実Excel が 何を 書いたか だけ 写します★（判じは 人が 紙を 見て する）
 *
 *  ★門★
 *    ①2本とも 無ければ 走らない（exit 6）
 *    ②式が 在る マスが 1つも 無ければ 走らない（exit 3）
 *
 *  使い方: node docs/measured/yomu-jitsu-excel-no-shirushi.mjs
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..');
const 出 = path.join(ここ, 'golden-jitsu-excel-no-shirushi-2026-09-20.tsv');
const require_ = createRequire(path.join(ROOT, 'package.json'));
/* ★zip を ほどくのは 借り物（SheetJS）の 中の 物を 使わず 手で 読みます★
   ＝`AdmZip` 等を 増やしたく ない ので ★包みの 中の 1本だけ★ を 自力で 取り出します */

/** ★zip の 中の 1本を 取り出す★（★store と deflate の 両方★） */
function zipから取る(buf, 欲しい) {
  /* ★中央目録を 後ろから 探す★（EOCD = 0x06054b50） */
  let eocd = -1;
  for (let i = buf.length - 22; i >= 0; i -= 1) {
    if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) return null;
  const 本数 = buf.readUInt16LE(eocd + 10);
  let p = buf.readUInt32LE(eocd + 16);
  for (let k = 0; k < 本数; k += 1) {
    if (buf.readUInt32LE(p) !== 0x02014b50) return null;
    const 詰め方 = buf.readUInt16LE(p + 10);
    const 圧んだ = buf.readUInt32LE(p + 20);
    const 名長 = buf.readUInt16LE(p + 28);
    const 追長 = buf.readUInt16LE(p + 30);
    const 注長 = buf.readUInt16LE(p + 32);
    const 頭位置 = buf.readUInt32LE(p + 42);
    const 名 = buf.slice(p + 46, p + 46 + 名長).toString('utf8');
    if (名 === 欲しい) {
      const 名長2 = buf.readUInt16LE(頭位置 + 26);
      const 追長2 = buf.readUInt16LE(頭位置 + 28);
      const 中身 = buf.slice(頭位置 + 30 + 名長2 + 追長2, 頭位置 + 30 + 名長2 + 追長2 + 圧んだ);
      if (詰め方 === 0) return 中身;
      const zlib = require_('node:zlib');
      return zlib.inflateRawSync(中身);
    }
    p += 46 + 名長 + 追長 + 注長;
  }
  return null;
}

const 読む物 = [
  { 札: '51(SEQUENCE ほか 7本)', 名: 'exally-jitsu-excel-kobore.xlsx' },
  { 札: '52(名簿に無い 11本)', 名: 'exally-jitsu-excel-meibo11.xlsx' },
  /* ★★2026-09-20 足し ── 55（分母から 漏れて いた 45個の 残り 21個）★★
     ＝Exally1 が 「09-16 の 分母 396個に ★45個 漏れて いた★」と 数えた 残り
     ＝★21個 とも 台が 持って いる＝お客さんが 打てます★
     ⇒★この 紙が そのまま 門（tests/xlfn-morenashi.test.mjs）の 分母に なります★
       ＝★足した 分だけ 門が 見る 範囲も 増えます★（Exally1 が 偽の 1行で 実測済み） */
  { 札: '55(残り 21本)', 名: 'exally-jitsu-excel-nokori21.xlsx' },
  /* ★★2026-09-20 足し ── 58（名簿に 「在る」16個＝★付け過ぎ★を 探す）★★
     ＝今まで 探して いたのは ★付け忘れ★（WRAPROWS / WRAPCOLS / MODE.MULT）
     ＝★逆も 壊れます★＝`_xlfn.` を 付け過ぎると 実Excel が 読めない（TRANSPOSE で 実測）
     ⇒★名簿に 在る 物が 本当に 要るかを 聞いた 事が 在りませんでした★ */
  { 札: '58(名簿に在る 16本)', 名: 'exally-jitsu-excel-meibo16.xlsx' },
  /* ★★2026-09-20 足し ── 59（抜き取り 6個＝まだ 聞いて いない 324個を 割る）★★
     ＝★FORMULATEXT が 一番 大事★（台が 持つ／名簿に 無い／2013年に 見える）
     ＝★他の 5個は 両端を 押さえる 為★ */
  { 札: '59(抜き取り 6本)', 名: 'exally-jitsu-excel-nukitori6.xlsx' },
  /* ★★2026-09-20 足し ── 61（★総当たり 314個★）★★
     ＝台が 持つ 376個の うち まだ 聞いて いなかった 319個を 全部 打たせた
     ＝★見立てで 絞るのを やめました★（穴 4個とも 「そこに 無い はず」の 所から 出た）
     ⇒★打てた 314個／打てない 5個★（打てない 物は 打ち方の 紙に 名前が 在ります）
     ⇒★この 紙が そのまま 門の 分母に なります★ */
  { 札: '61(総当たり 314本)', 名: 'exally-jitsu-excel-sounate.xlsx' },
];

const 行 = [];
行.push('# ★実Excel 自身が 溢れる 式に 付ける 印★（51・52）（2026-09-20）');
行.push('# ★読むだけ★（1バイトも 書いて いません）');
行.push('# ★作った 物★ ... docs/measured/tsukuru-jitsu-excel-kobore-6shurui.ps1');
行.push('#                docs/measured/tsukuru-jitsu-excel-meibo-ni-nai-11.ps1');
行.push('# ★実Excel★ ... 版 16.0 ／ build 20326（作った 道具が 出して います）');
行.push('# ★判じは ここでは しません★＝実Excel が 書いた 字を そのまま 写すだけ');

let 全部 = 0;
for (const x of 読む物) {
  const p = path.join(os.tmpdir(), x.名);
  if (!fs.existsSync(p)) {
    console.log('★★在りません ... ' + p + '★★');
    process.exit(6);
  }
  const buf = fs.readFileSync(p);
  const xml = zipから取る(buf, 'xl/worksheets/sheet1.xml');
  if (!xml) { console.log('★★sheet1.xml を 取り出せません ... ' + x.名 + '★★'); process.exit(6); }
  const s = xml.toString('utf8');
  行.push('#');
  行.push('# ═══ ★' + x.札 + '★ ... ' + x.名 + '（' + buf.length + ' バイト）═══');
  行.push('# ★cm= の 数★ ' + (s.match(/cm=/g) || []).length + ' ／ ★t="array" の 数★ ' + (s.match(/t="array"/g) || []).length);
  行.push('# マス\tcm\tref\t実Excel が 書いた 式');
  const re = /<c r="([A-Z]+\d+)"([^>]*)>\s*<f([^>]*)>([\s\S]*?)<\/f>/g;
  let m;
  let 本数 = 0;
  while ((m = re.exec(s)) !== null) {
    const マス = m[1];
    const cm = /cm="/.test(m[2]) ? '在り' : '無し';
    const ref = (/ref="([^"]+)"/.exec(m[3]) || [null, '(無し)'])[1];
    const f = m[4].replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&').replace(/&quot;/g, '"');
    行.push(マス + '\t' + cm + '\t' + ref + '\t' + f);
    本数 += 1;
    全部 += 1;
  }
  行.push('# ★式が 在る マス ... ' + 本数 + '個★');
  console.log('  ' + x.札 + ' ... 式が 在る マス ' + 本数 + '個');
}

if (全部 === 0) { console.log('★★式が 1つも 在りません★★'); process.exit(3); }

fs.writeFileSync(出, 行.join('\n') + '\n', 'utf8');
console.log('');
console.log('★書いた ... ' + 出 + '★（★' + 全部 + '個★）');
console.log('');
console.log('★★断り★★');
console.log('  ・★実Excel が 書いた 字を 写すだけ＝「うちが 正しいか」は 判じて いません★');
console.log('  ・★1つの 版（16.0 build 20326）だけ★');
console.log('  ・★日本語の Excel です★（国で 別名が 変わる 物が 在ります）');
