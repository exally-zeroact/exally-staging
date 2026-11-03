/* sakusha-mihon.test.mjs -- ★Office ファイルの 作者の 欄に 名前が 戻って いないか★（2026-10-10）
 *
 *  ★なぜ★ 公開 repo の 見本・測った 証しの ファイル 18本の 作者の 欄に、作った 人の 名前が 入って いた
 *    （司さんの 決め 10-10「実在の 名前は 架空に」＝scripts/sakusha-mihon.mjs で「見本」に 替えた）。
 *    Excel は 保存の たびに サインインの 名前を 入れる ので、作り直した 日に 黙って 戻る。
 *  ★判じ★ … 白名簿＝作者 2欄が「見本」か 欄なし だけを 通す（名前は ここに 書かない）。
 *    分母は git ls-files（置き場で 決め打ち しない）。読めない zip・zip64・データ記述子は 赤。
 *  --self-test … 作り物の 名前を 入れた 写し・読めない zip が 赤に なるか
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import zlib from 'node:zlib';
import { fileURLToPath } from 'node:url';
import { 作者を読む, 作者を見本に, 見本にする, 欄の判じ, Officeファイル, 見本 } from '../scripts/sakusha-mihon.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
let 緑 = 0, 赤 = 0;
const T = (n, ok, m) => { if (ok) { 緑++; console.log('  ok   ' + n); } else { 赤++; console.log('  NG   ' + n + (m ? '\n       ' + m : '')); } };
const 判じる = (buf) => { try { return 欄の判じ(作者を読む(buf)); } catch (e) { return '読めない'; } };

if (process.argv.includes('--self-test')) {
  console.log('[sakusha-mihon --self-test] ★名前が 戻ると 赤に なるか★');
  const 見本の本 = Officeファイル(ROOT).find((f) => 判じる(fs.readFileSync(path.join(ROOT, f))) === '見本');
  T('★見本の 本が 1本 以上 在る★', !!見本の本);
  const 元 = fs.readFileSync(path.join(ROOT, 見本の本));
  /* ①作者の 字を 作り物の 名前に した 写し ⇒ その他（赤） */
  const 置き場 = fs.mkdtempSync(path.join(os.tmpdir(), 'sakusha-self-'));
  try {
    const 写し = 書き換え(元, '試験太郎');
    T('★作者が 作り物の 名前の 写しは「その他」（赤）★', 判じる(写し) === 'その他', 判じる(写し));
    T('★その 写しを 道具に 通すと「見本」に 戻る★', 判じる(作者を見本に(写し)) === '見本');
    /* ④拾い損ねる 書き方（10-10 本番前の 対立役＝前の 判じでは 全部 緑だった）⇒ どれも その他（赤） */
    const 名 = '試験太郎';
    const 形 = [
      ['閉じタグに 空白', (t) => t.replace('>' + 見本 + '</dc:creator>', '>' + 名 + '</dc:creator >')],
      ['CDATA', (t) => t.replace('<dc:creator>' + 見本 + '</dc:creator>', '<dc:creator><![CDATA[' + 名 + ']]></dc:creator>')],
      ['2つ目の creator', (t) => t.replace('</cp:coreProperties>', '<dc:creator>' + 名 + '</dc:creator></cp:coreProperties>')],
      ['接頭辞 違い', (t) => t.replace('</cp:coreProperties>', '<ns0:creator xmlns:ns0="http://purl.org/dc/elements/1.1/">' + 名 + '</ns0:creator></cp:coreProperties>')],
      ['既定の 名前空間', (t) => t.replace('</cp:coreProperties>', '<creator xmlns="http://purl.org/dc/elements/1.1/">' + 名 + '</creator></cp:coreProperties>')],
      ['拾えない 形（属性の 中に >）', (t) => t.replace('</cp:coreProperties>', '<dc:creator a=">">' + 名 + '</dc:creator></cp:coreProperties>')],
    ];
    for (const [n, fn] of 形) {
      const w = 書き換え(元, '', 'docProps/core.xml', fn);
      T('★' + n + ' は 赤★', 判じる(w) !== '見本' && 判じる(w) !== '欄なし', 判じる(w));
    }
    const 会社 = 書き換え(元, '', 'docProps/app.xml', (t) => t.indexOf('<Company/>') >= 0 ? t.replace('<Company/>', '<Company>' + 名 + '</Company>')
      : t.indexOf('<Company>') >= 0 ? t.replace(/<Company>[^<]*<\/Company>/, '<Company>' + 名 + '</Company>')
      : t.replace('</Properties>', '<Company>' + 名 + '</Company></Properties>'));
    T('★app.xml の Company に 字 は 赤★', 判じる(会社) === 'その他', 判じる(会社));
    /* ⑤保存した PC の 道（workbook.xml の absPath）に 字 ⇒ 赤（10-11 禁止の 字の 見張りで 見つけた） */
    const 作り物の道 = 'C:/試験/見本/';
    const 道 = 書き換え(元, '', 'xl/workbook.xml', (t) => /absPath\b[^>]*\burl="/.test(t)
      ? t.replace(/(absPath\b[^>]*?\burl=")[^"]*"/, '$1' + 作り物の道 + '"')
      : t.replace('</workbook>', '<x15ac:absPath xmlns:x15ac="http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac" url="' + 作り物の道 + '"/></workbook>'));
    T('★保存した PC の 道に 字 は 赤★', 判じる(道) === 'その他', 判じる(道));
    T('★その 写しを 道具に 通すと 道が 空に 戻る★', 判じる(見本にする(道)) === '見本');
    /* ⑥道の 書き方の 揺れ（10-11 本番前の 対立役＝前の 判じでは 緑） */
    const 空白 = 書き換え(元, '', 'xl/workbook.xml', (t) => t.replace(/(absPath\b[^>]*?\burl)="[^"]*"/, '$1 = "' + 作り物の道 + '"'));
    T('★url = "道"（= の 前後に 空白）の 道 は 赤★', 判じる(空白) === 'その他', 判じる(空白));
    const 単 = 書き換え(元, '', 'xl/workbook.xml', (t) => t.replace(/(absPath\b[^>]*?\burl)="[^"]*"/, "$1='" + 作り物の道 + "'"));
    T("★url='道'（単引用符）の 道 は 赤★", 判じる(単) === 'その他', 判じる(単));
    T('★単引用符の 道も 道具で 空に 戻る★', 判じる(見本にする(単)) === '見本');
    /* ⑦xlsb の 道（workbook.bin の 記録 0x817） */
    const xlsb本 = Officeファイル(ROOT).find((p) => /\.xlsb$/i.test(p));
    T('★xlsb の 見本が 在る★', !!xlsb本);
    const xb = xlsbに道を入れる(fs.readFileSync(path.join(ROOT, xlsb本)), 作り物の道);
    T('★xlsb の 道の 記録に 字 は 赤★', 判じる(xb) === 'その他', 判じる(xb));
    T('★xlsb の 道も 道具で 空に 戻る★', 判じる(見本にする(xb)) === '見本');
    /* ②読めない zip ⇒ 読めない（赤） */
    T('★読めない zip は「読めない」（赤）★', 判じる(Buffer.from('PK\u0003\u0004 こわれ')) === '読めない');
    /* ③データ記述子の 印を 立てた zip ⇒ 読めない（赤） */
    const 印 = Buffer.from(元);
    const e = 印.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
    const cd = 印.readUInt32LE(e + 16);
    印.writeUInt16LE(印.readUInt16LE(cd + 8) | 8, cd + 8);
    T('★データ記述子の zip は「読めない」（赤）★', 判じる(印) === '読めない');
  } finally { fs.rmSync(置き場, { recursive: true, force: true }); }
  console.log('\nsakusha-mihon --self-test: ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
  process.exit(赤 ? 1 : 0);
}

/** xlsb の workbook.bin の 道の 記録（0x817）に 字を 入れた 写し（自己試験用） */
function xlsbに道を入れる(buf, 道) {
  return 書き換えbin(buf, 'xl/workbook.bin', (bin) => {
    const 片 = []; let p = 0; let 入れた = false;
    while (p < bin.length) {
      const 頭 = p; let t = 0, k = 0, c;
      do { c = bin[p++]; t |= (c & 0x7f) << (7 * k++); } while ((c & 0x80) && k < 2);
      const 種の終 = p; let l = 0; k = 0;
      do { c = bin[p++]; l += (c & 0x7f) * Math.pow(2, 7 * k++); } while ((c & 0x80) && k < 4);
      if (t === 0x817) {
        const 字 = Buffer.from(道, 'utf16le'); const 中 = Buffer.alloc(4 + 字.length); 中.writeUInt32LE(道.length, 0); 字.copy(中, 4);
        const 長 = []; let n = 中.length; do { let x = n % 128; n = Math.floor(n / 128); if (n) x |= 0x80; 長.push(x); } while (n);
        片.push(bin.subarray(頭, 種の終), Buffer.from(長), 中); 入れた = true;
      } else 片.push(bin.subarray(頭, p + l));
      p += l;
    }
    if (!入れた) throw new Error('xlsb に 道の 記録が 無い');
    return Buffer.concat(片);
  });
}
/** 部品を 生の バイトで 差し替えた 写し */
function 書き換えbin(buf, 部品名, 替え) {
  return 書き換え(buf, '', 部品名, null, 替え);
}

/** 作者の 2欄を 字に 替えた 写し（道具の 書き方を 使い、最後に 字だけ 差し替える） */
function 書き換え(buf, 字, 部品名, 替え, バイトの替え) {
  /* 道具は「見本」に しか 替えない ので、core.xml を 開いて 字を 入れ、同じ 方式で 詰め直す */
  const b = Buffer.from(buf);
  const e = b.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]));
  const 数 = b.readUInt16LE(e + 10);
  let p = b.readUInt32LE(e + 16);
  const 部品 = [];
  for (let k = 0; k < 数; k++) {
    const nl = b.readUInt16LE(p + 28), xl = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
    部品.push({ 目録: p, 名: b.toString('utf8', p + 46, p + 46 + nl), 位置: b.readUInt32LE(p + 42), 圧縮: b.readUInt32LE(p + 20), 方式: b.readUInt16LE(p + 10) });
    p += 46 + nl + xl + cl;
  }
  const x = 部品.find((q) => q.名 === (部品名 || 'docProps/core.xml'));
  const lnl = b.readUInt16LE(x.位置 + 26), lxl = b.readUInt16LE(x.位置 + 28), 頭 = 30 + lnl + lxl;
  const raw = b.subarray(x.位置 + 頭, x.位置 + 頭 + x.圧縮);
  const 中 = (x.方式 === 8 ? zlib.inflateRawSync(raw) : raw).toString('utf8');
  const 中2 = 替え ? 替え(中) : 中.split(見本).join(字);
  const 生 = バイトの替え ? バイトの替え(x.方式 === 8 ? zlib.inflateRawSync(raw) : Buffer.from(raw)) : Buffer.from(中2, 'utf8');
  const 新 = x.方式 === 8 ? zlib.deflateRawSync(生) : 生;
  const 差 = 新.length - raw.length;
  const h = Buffer.from(b.subarray(x.位置, x.位置 + 頭));
  h.writeUInt32LE(zlib.crc32(生) >>> 0, 14); h.writeUInt32LE(新.length, 18); h.writeUInt32LE(生.length, 22);
  const 前 = b.subarray(0, x.位置), 後 = b.subarray(x.位置 + 頭 + x.圧縮, e);
  const 目録の位置 = b.readUInt32LE(e + 16);
  const out = Buffer.concat([前, h, 新, 後, b.subarray(e)]);
  /* 目録と 終わりの 位置を 直す */
  const 新目録 = 目録の位置 + 差;
  const 新e = e + 差;
  out.writeUInt32LE(新目録, 新e + 16);
  for (const q of 部品) {
    const o = q.目録 + 差;
    if (q === x) { out.writeUInt32LE(zlib.crc32(生) >>> 0, o + 16); out.writeUInt32LE(新.length, o + 20); out.writeUInt32LE(生.length, o + 24); }
    if (q.位置 > x.位置) out.writeUInt32LE(q.位置 + 差, o + 42);
  }
  return out;
}

console.log('[sakusha-mihon] ★Office ファイルの 作者の 欄は「見本」か 欄なし だけ★');
const 本 = Officeファイル(ROOT);
const 数 = { 見本: 0, 欄なし: 0, その他: 0, 読めない: 0 };
const 外 = [];
for (const f of 本) { const k = 判じる(fs.readFileSync(path.join(ROOT, f))); 数[k]++; if (k === 'その他' || k === '読めない') 外.push(k + ' ' + f); }
T('★Office ファイルを 実際に 数えて いる（' + 本.length + '本）★', 本.length >= 20, String(本.length));
T('★作者の 欄は「見本」か 欄なし だけ（見本 ' + 数.見本 + '・欄なし ' + 数.欄なし + '・その他 ' + 数.その他 + '・読めない ' + 数.読めない + '）★',
  数.その他 === 0 && 数.読めない === 0, 外.join('\n       '));
console.log('\nsakusha-mihon: ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
process.exit(赤 ? 1 : 0);
