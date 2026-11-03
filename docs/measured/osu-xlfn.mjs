/* osu-xlfn.mjs — ★実Excel が 保存した ファイルの 中で 式を どう 書くか★ を 読む（2026-09-16）
 *
 *  ★★なぜ★★
 *    うちが 書いた `=NORM.DIST(...)` を 実Excel が ★#NAME?★ に する。
 *    でも ★実Excel 自身が 打つと ちゃんと 答える★（0.841344746068543）。
 *    ⇒★関数を 知らないのでは なく ★ファイルの 中の 書き方★が 違う★
 *    ⇒★当てずに 読む★＝実Excel が 保存した 物を そのまま 開く
 *
 *  使い方: node docs/measured/osu-xlfn.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { ROOT } from './honban-no-michi.mjs';

const 元 = path.join(ROOT, 'docs/measured/xlfn-shirabe.xlsx');
const うち = path.join(ROOT, 'docs/measured/oufuku-shinki.xlsx');

/* ★zip を 解く★（★中の 1つの 名前だけ★） */
function zipから(ファイル, 欲しい) {
  const b = fs.readFileSync(ファイル);
  /* ★末尾の 中央目録から 辿る★ */
  let eocd = -1;
  for (let i = b.length - 22; i >= 0 && i > b.length - 70000; i--) {
    if (b.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) throw new Error('★zip の 終わりが 見つからない★');
  const 数 = b.readUInt16LE(eocd + 10);
  let p = b.readUInt32LE(eocd + 16);
  for (let k = 0; k < 数; k++) {
    if (b.readUInt32LE(p) !== 0x02014b50) break;
    const 法 = b.readUInt16LE(p + 10);
    const 圧 = b.readUInt32LE(p + 20);
    const 名長 = b.readUInt16LE(p + 28);
    const 追長 = b.readUInt16LE(p + 30);
    const 注長 = b.readUInt16LE(p + 32);
    const 先 = b.readUInt32LE(p + 42);
    const 名 = b.toString('utf8', p + 46, p + 46 + 名長);
    if (名 === 欲しい) {
      const ln = b.readUInt16LE(先 + 26), le = b.readUInt16LE(先 + 28);
      const 中 = b.slice(先 + 30 + ln + le, 先 + 30 + ln + le + 圧);
      return 法 === 0 ? 中 : zlib.inflateRawSync(中);
    }
    p += 46 + 名長 + 追長 + 注長;
  }
  throw new Error('★中に ' + 欲しい + ' が 無い★');
}

function 式たち(ファイル) {
  const xml = zipから(ファイル, 'xl/worksheets/sheet1.xml').toString('utf8');
  const 出 = [];
  const re = /<c r="([A-Z]+\d+)"[^>]*>(?:(?!<\/c>).)*?<f[^>]*>([^<]*)<\/f>/g;
  let m;
  while ((m = re.exec(xml)) !== null) 出.push([m[1], m[2]]);
  return 出;
}

console.log('');
console.log('★★実Excel が 保存した ファイルの 中★★（xl/worksheets/sheet1.xml）');
const 実 = 式たち(元);
for (const [マス, f] of 実) {
  const 印 = f.indexOf('_xlfn.') >= 0 ? '★_xlfn. が 付く★' : '';
  console.log('  ' + マス.padEnd(5) + f.padEnd(44) + 印);
}

console.log('');
console.log('★★うちが 書き出した ファイルの 中★★');
const 我 = 式たち(うち);
for (const [マス, f] of 我) {
  const 印 = f.indexOf('_xlfn.') >= 0 ? '★_xlfn. が 付く★' : '';
  console.log('  ' + マス.padEnd(5) + f.padEnd(44) + 印);
}

console.log('');
const 付く = 実.filter(([, f]) => f.indexOf('_xlfn.') >= 0).map(([, f]) => f.replace(/^_xlfn\./, '').replace(/\(.*$/, ''));
const 付かない = 実.filter(([, f]) => f.indexOf('_xlfn.') < 0).map(([, f]) => f.replace(/^=/, '').replace(/\(.*$/, ''));
console.log('★★実Excel が `_xlfn.` を 付ける 関数★★ … ' + 付く.join(' '));
console.log('★実Excel が 付けない 関数★ … ' + 付かない.join(' '));
