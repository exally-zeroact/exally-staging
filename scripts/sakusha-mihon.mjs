/* sakusha-mihon.mjs -- ★Office ファイル（xlsx/xlsb/xlsm）の 作者の 欄を「見本」に 替える★（2026-10-10）
 *
 *  ★なぜ★ 公開 repo の 見本・測った 証しの ファイルの docProps/core.xml に、作った 人の 名前が 入って いた
 *    （司さんの 決め 10-10「実在の 名前は 架空に」）。Excel は サインインの アカウント名を 入れるので、
 *    docs/measured の toru-*.ps1 4本と tools/make-mihon*.ps1 は 保存の 後に この 道具を 呼ぶ（判じは ここ 1か所）。
 *    ほかの 見本の 作り手は 呼ばない＝作り直して 名前が 戻れば 門（tests/sakusha-mihon.test.mjs）が 赤で 拾う。
 *  ★替える 所★ … docProps/core.xml の dc:creator と cp:lastModifiedBy の 字だけ（欄が 無い 物は 足さない）。
 *    ほかの 部品は ★圧縮の バイトまで そのまま★（並び・拡張の 欄も 保つ）。後ろの 部品の 位置だけ ずらす。
 *  ★止まる 時（赤）★ … zip が 読めない／zip64／データ記述子（flag bit3）／core.xml が UTF-16／欄の 書き方が 拾えない／
 *    app.xml の Company・Manager に 字／替えた 後に「見本」に ならない
 *  使い方：
 *    node scripts/sakusha-mihon.mjs <ファイル>...      … 替える（もう 見本なら 何もしない）
 *    node scripts/sakusha-mihon.mjs --check            … git ls-files の Office ファイルを 全部 見る（替えない）
 *  ★名前は 出さない★（出すのは「見本／欄なし／その他」だけ）
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

export const 見本 = '見本';
const CORE = 'docProps/core.xml';

function 目録を読む(b) {
  let e = -1;
  for (let i = b.length - 22; i >= Math.max(0, b.length - 22 - 65535); i--) if (b.readUInt32LE(i) === 0x06054b50) { e = i; break; }
  if (e < 0) throw new Error('zip の 終わりが 無い');
  const 数 = b.readUInt16LE(e + 10), 目録の大きさ = b.readUInt32LE(e + 12), 目録の位置 = b.readUInt32LE(e + 16);
  if (数 === 0xffff || 目録の位置 === 0xffffffff) throw new Error('zip64 は 扱わない');
  const 部品 = [];
  let p = 目録の位置;
  for (let k = 0; k < 数; k++) {
    if (b.readUInt32LE(p) !== 0x02014b50) throw new Error('目録が 壊れて いる');
    const nl = b.readUInt16LE(p + 28), xl = b.readUInt16LE(p + 30), cl = b.readUInt16LE(p + 32);
    const flag = b.readUInt16LE(p + 8);
    if (flag & 8) throw new Error('データ記述子は 扱わない');
    部品.push({ 目録: p, 目録の長さ: 46 + nl + xl + cl, 名: b.toString('utf8', p + 46, p + 46 + nl), 位置: b.readUInt32LE(p + 42), 圧縮: b.readUInt32LE(p + 20), 方式: b.readUInt16LE(p + 10) });
    p += 46 + nl + xl + cl;
  }
  return { 終わり: e, 目録の位置, 目録の大きさ, 部品 };
}

function 部品の中身(b, x) {
  if (b.readUInt32LE(x.位置) !== 0x04034b50) throw new Error('部品の 頭が 壊れて いる');
  const nl = b.readUInt16LE(x.位置 + 26), xl = b.readUInt16LE(x.位置 + 28);
  const 頭 = 30 + nl + xl;
  const raw = b.subarray(x.位置 + 頭, x.位置 + 頭 + x.圧縮);
  return { 頭, raw, 中: x.方式 === 8 ? zlib.inflateRawSync(raw) : x.方式 === 0 ? Buffer.from(raw) : null };
}

/* ★欄を 拾う 形★＝接頭辞（dc: 等・違っても 無くても）・属性・閉じタグの 空白・自己閉じ・CDATA を 拾う
 *  （10-10 本番前の 対立役＝閉じタグの 空白・CDATA・2つ目の 欄・接頭辞違い・既定の 名前空間で 素通りした） */
const 欄の形 = /<(?:[\w.-]+:)?(creator|lastModifiedBy)\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?\1\s*>)/gi;
const 中身の字 = (s) => String(s === undefined ? '' : s).replace(/^<!\[CDATA\[([\s\S]*)\]\]>$/, '$1');

function 部品を探す(d, 名) { return d.部品.find((p) => p.名.toLowerCase() === 名.toLowerCase()); }
function 字にする(buf) {
  if (buf.length >= 2 && ((buf[0] === 0xff && buf[1] === 0xfe) || (buf[0] === 0xfe && buf[1] === 0xff))) return { 字: null, utf16: true };
  return { 字: buf.toString('utf8'), utf16: false };
}

/** 作者の 欄を 読む ⇒ null（core.xml 無し）／{ 値: [拾えた 欄の 字], 残り: 拾えなかった 欄の 字が 在るか, 会社: [app.xml の Company・Manager] } */
export function 作者を読む(buf) {
  const b = Buffer.from(buf);
  const d = 目録を読む(b);
  const x = 部品を探す(d, CORE);
  if (!x) { const 道 = 道を読む(b); return 道.length ? { 値: [], 残り: false, 会社: [], 道 } : null; }
  const s = 部品の中身(b, x).中;
  if (!s) throw new Error('core.xml の 圧縮の 方式が 読めない');
  const { 字, utf16 } = 字にする(s);
  if (utf16) throw new Error('core.xml が UTF-16（読まない＝赤）');
  const 値 = [];
  const 残りの字 = 字.replace(欄の形, (m, 名, 中) => { 値.push(中身の字(中)); return ''; });
  const 会社 = [];
  const app = 部品を探す(d, 'docProps/app.xml');
  if (app) {
    const a = 部品の中身(b, app).中;
    const t = a ? a.toString('utf8') : '';
    for (const m of t.matchAll(/<(?:[\w.-]+:)?(Company|Manager)\b[^>]*?(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?\1\s*>)/gi)) 会社.push(中身の字(m[2]));
  }
  return { 値, 残り: /creator|lastModifiedBy/i.test(残りの字), 会社, 道: 道を読む(b) };
}

/* ★保存した PC の 道★（利用者の フォルダ名が 入る・10-11 禁止の 字の 見張り）
 *   xlsx/xlsm ＝ xl/workbook.xml の absPath の url（= の 前後の 空白・単引用符 も 拾う。UTF-16 の workbook.xml は 赤）
 *   xlsb      ＝ xl/workbook.bin の 記録 BrtAbsPath15（種類 0x817・字は XLWideString＝4バイトの 字数＋UTF-16LE）
 *   （10-11 本番前の 対立役＝xlsb 8本に 道が 残り、前の 道具は 緑を 返して いた） */
const 道の形 = /(<(?:[\w.-]+:)?absPath\b[^>]*?\burl\s*=\s*)(["'])([^"']*)\2/gi;
const ABS_PATH = 0x817;
/** BIFF12 の 記録を 並べる ⇒ [{ 種類, 頭, 長さ, 中の位置 }]（壊れて いれば 投げる） */
function 記録たち(buf) {
  const 出 = [];
  let p = 0;
  while (p < buf.length) {
    const 頭 = p;
    let 種類 = 0, k = 0, c;
    do { if (p >= buf.length) throw new Error('xlsb の 記録が 途中で 切れた'); c = buf[p++]; 種類 |= (c & 0x7f) << (7 * k++); } while ((c & 0x80) && k < 2);
    let 長さ = 0; k = 0;
    do { if (p >= buf.length) throw new Error('xlsb の 記録が 途中で 切れた'); c = buf[p++]; 長さ += (c & 0x7f) * Math.pow(2, 7 * k++); } while ((c & 0x80) && k < 4);
    if (p + 長さ > buf.length) throw new Error('xlsb の 記録が 途中で 切れた');
    出.push({ 種類, 頭, 中の位置: p, 長さ });
    p += 長さ;
  }
  return 出;
}
function 長さの字(n) { const o = []; do { let c = n % 128; n = Math.floor(n / 128); if (n) c |= 0x80; o.push(c); } while (n); return Buffer.from(o); }
function 道の部品(d) { return 部品を探す(d, 'xl/workbook.xml') || 部品を探す(d, 'xl/workbook.bin'); }
export function 道を読む(buf) {
  const b = Buffer.from(buf);
  const d = 目録を読む(b);
  const x = 道の部品(d);
  if (!x) return [];
  const 中 = 部品の中身(b, x).中;
  if (!中) throw new Error('workbook の 圧縮の 方式が 読めない');
  if (/\.bin$/i.test(x.名)) {
    return 記録たち(中).filter((r) => r.種類 === ABS_PATH).map((r) => {
      const n = r.長さ >= 4 ? 中.readUInt32LE(r.中の位置) : 0;
      return 中.subarray(r.中の位置 + 4, r.中の位置 + 4 + n * 2).toString('utf16le');
    });
  }
  const { 字, utf16 } = 字にする(中);
  if (utf16) throw new Error('workbook.xml が UTF-16（読まない＝赤）');
  const 値 = [...字.matchAll(道の形)].map((m) => m[3]);
  /* 拾えない 書き方で absPath が 在れば 赤（拾えた 数と 字の 数が 合わない） */
  if ((字.match(/absPath\b/gi) || []).length !== 値.length) 値.push('★拾えない 書き方★');
  return 値;
}
/** 保存した PC の 道を 空に ⇒ 新しい Buffer（無ければ 同じ物） */
export function 道を消す(buf) {
  const b = Buffer.from(buf);
  const d = 目録を読む(b);
  const x = 道の部品(d);
  if (!x) return b;
  const 中 = 部品の中身(b, x).中;
  if (!中) throw new Error('workbook の 圧縮の 方式が 読めない');
  let 後;
  if (/\.bin$/i.test(x.名)) {
    const 片 = [];
    let 替えた = false;
    for (const r of 記録たち(中)) {
      const 元 = 中.subarray(r.頭, r.中の位置 + r.長さ);
      if (r.種類 === ABS_PATH && !(r.長さ >= 4 && 中.readUInt32LE(r.中の位置) === 0)) {
        const 種類の字 = 中.subarray(r.頭, r.頭 + (r.種類 >= 0x80 ? 2 : 1));
        片.push(種類の字, 長さの字(4), Buffer.alloc(4));
        替えた = true;
      } else 片.push(元);
    }
    if (!替えた) return b;
    後 = Buffer.concat(片);
  } else {
    const { 字: 前, utf16 } = 字にする(中);
    if (utf16) throw new Error('workbook.xml が UTF-16（替えない＝赤）');
    const 替 = 前.replace(道の形, '$1$2$2');
    if (替 === 前) return b;
    後 = Buffer.from(替, 'utf8');
  }
  const out = 部品を差し替える(b, d, x, 後);
  if (道を読む(out).some((v) => v !== '')) throw new Error('道を 空に できない');
  return out;
}
/** 作者の 欄を 見本に し、保存した PC の 道を 空に する */
export function 見本にする(buf) { return 道を消す(作者を見本に(buf)); }

/** 判じ：見本 か 欄なし だけを 通す（拾えない 書き方・空で ない 会社は その他＝赤） */
export function 欄の判じ(r) {
  if (r === null) return '欄なし';
  if (r.残り) return 'その他';
  if (r.値.some((v) => v !== '' && v !== 見本)) return 'その他';
  if ((r.会社 || []).some((v) => v !== '')) return 'その他';
  if ((r.道 || []).some((v) => v !== '')) return 'その他';
  return r.値.some((v) => v === 見本) ? '見本' : '欄なし';
}

/** 部品 x の 中身を 生 に 差し替えた zip を 作る（ほかの 部品は 圧縮の バイトまで そのまま・後ろの 位置だけ ずらす） */
function 部品を差し替える(b, d, x, 生) {
  const { 頭, raw } = 部品の中身(b, x);
  const 新raw = x.方式 === 8 ? zlib.deflateRawSync(生, { level: 9 }) : 生;
  const crc = zlib.crc32(生) >>> 0;
  const 差 = 新raw.length - raw.length;
  /* 部品を 位置の 順に 並べ、x だけ 差し替えて つなぐ（すき間が 在れば 止める） */
  const 順 = d.部品.slice().sort((a, c) => a.位置 - c.位置);
  const 片 = [];
  let 次 = 0;
  for (const p of 順) {
    if (p.位置 !== 次) throw new Error('部品の 間に すき間が 在る');
    const 中身 = 部品の中身(b, p);
    const 終 = p.位置 + 中身.頭 + p.圧縮;
    if (p === x) {
      const h = Buffer.from(b.subarray(p.位置, p.位置 + 頭));
      h.writeUInt32LE(crc, 14); h.writeUInt32LE(新raw.length, 18); h.writeUInt32LE(生.length, 22);
      片.push(h, 新raw);
    } else 片.push(b.subarray(p.位置, 終));
    次 = 終;
  }
  if (次 !== d.目録の位置) throw new Error('部品と 目録の 間に すき間が 在る');
  const 目録 = Buffer.from(b.subarray(d.目録の位置, d.目録の位置 + d.目録の大きさ));
  for (const p of d.部品) {
    const o = p.目録 - d.目録の位置;
    if (p === x) { 目録.writeUInt32LE(crc, o + 16); 目録.writeUInt32LE(新raw.length, o + 20); 目録.writeUInt32LE(生.length, o + 24); }
    if (p.位置 > x.位置) 目録.writeUInt32LE(p.位置 + 差, o + 42);
  }
  const 終わり = Buffer.from(b.subarray(d.終わり));
  終わり.writeUInt32LE(d.目録の位置 + 差, 16);
  const out = Buffer.concat([...片, 目録, b.subarray(d.目録の位置 + d.目録の大きさ, d.終わり), 終わり]);
  return out;
}

/** 作者の 欄を「見本」に ⇒ 新しい Buffer（替える 物が 無ければ 同じ物） */
export function 作者を見本に(buf) {
  const b = Buffer.from(buf);
  const d = 目録を読む(b);
  const x = 部品を探す(d, CORE);
  if (!x) return b;
  const { 頭, raw, 中 } = 部品の中身(b, x);
  if (!中) throw new Error('core.xml の 圧縮の 方式が 読めない');
  const { 字: 前, utf16 } = 字にする(中);
  if (utf16) throw new Error('core.xml が UTF-16（替えない＝赤）');
  const 後 = 前.replace(欄の形, (m, 名, 内) => (内 === undefined ? m : m.replace('>' + 内 + '<', '>' + 見本 + '<')));
  if (後 === 前) {
    if (欄の判じ(Object.assign(作者を読む(b), { 道: [] })) === 'その他') throw new Error('作者の 欄を 替えられない（書き方が 拾えない か 会社の 欄に 字）');
    return b;
  }
  const out = 部品を差し替える(b, d, x, Buffer.from(後, 'utf8'));
  if (欄の判じ(Object.assign(作者を読む(out), { 道: [] })) !== '見本') throw new Error('替えた 後も 作者の 欄が「見本」に ならない');
  return out;
}

export function Officeファイル(root) {
  return execFileSync('git', ['-C', root, 'ls-files'], { encoding: 'utf8' }).split('\n')
    .filter((f) => /\.(xlsx|xlsb|xlsm|docx|pptx)$/i.test(f));
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  process.exitCode = 1;
  const 引数 = process.argv.slice(2);
  const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
  if (引数[0] === '--check') {
    const 数 = { 見本: 0, 欄なし: 0, その他: 0, 読めない: 0 };
    for (const f of Officeファイル(ROOT)) {
      let k;
      try { k = 欄の判じ(作者を読む(fs.readFileSync(path.join(ROOT, f)))); } catch (e) { k = '読めない'; }
      数[k]++;
      if (k === 'その他' || k === '読めない') console.log('  ★' + k + '★ ' + f);
    }
    console.log('Office ' + Object.values(数).reduce((a, c) => a + c, 0) + '本：見本 ' + 数.見本 + '・欄なし ' + 数.欄なし + '・その他 ' + 数.その他 + '・読めない ' + 数.読めない);
    if (!数.その他 && !数.読めない) process.exitCode = 0;
  } else {
    if (!引数.length) { console.log('使い方: node scripts/sakusha-mihon.mjs <ファイル>... ／ --check'); }
    else {
      for (const f of 引数) {
        const 前 = fs.readFileSync(f);
        const 後 = 見本にする(前);
        if (後 !== 前 && !後.equals(前)) { fs.writeFileSync(f, 後); console.log('見本に 替えた … ' + f + '（' + 前.length + ' → ' + 後.length + ' バイト）'); }
        else console.log('替える 物 無し … ' + f);
      }
      process.exitCode = 0;
    }
  }
}
