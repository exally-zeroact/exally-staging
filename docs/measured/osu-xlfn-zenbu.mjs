/* osu-xlfn-zenbu.mjs — ★台が 知る 関数 全部★の `_xlfn.` を 測る（2026-09-16）
 *
 *  ★★なぜ★★
 *    `lib/xlsx-io.js` の 一覧に 漏れが 有ると ★その 式だけ #NAME?★ に なる。
 *    ★同じ 型で 既に 2回 踏んで います★（PERMUTATIONA ／ RANK.AVG）
 *    ⇒★手で 足すと また 漏れる★＝★台が 知る 物を 全部 実Excel に 聞く★
 *
 *  使い方:
 *    ① node docs/measured/osu-xlfn-zenbu.mjs --namae   … ★名簿を 書く★
 *    ② pwsh -NoProfile -File docs/measured/toru-xlfn-zenbu.ps1  … ★実Excel に 打たせる★
 *    ③ node docs/measured/osu-xlfn-zenbu.mjs           … ★中を 読んで 突き合わせる★
 */
import fs from 'node:fs';
import path from 'node:path';
import zlib from 'node:zlib';
import { createRequire } from 'node:module';
import { ROOT } from './honban-no-michi.mjs';

const require_ = createRequire(path.join(ROOT, 'package.json'));
const 名簿 = path.join(ROOT, 'docs/measured/xlfn-namae.txt');
const 元 = path.join(ROOT, 'docs/measured/xlfn-zenbu.xlsx');

const K = require_(path.join(ROOT, 'lib/shiki-kansuu.js'));

/* ★台が 知る 名前★（★手で 並べません★） */
const 台の名 = Object.keys(K.表).sort();

if (process.argv.indexOf('--namae') >= 0) {
  const 行 = ['# ★台が 知る 関数の 名簿★（osu-xlfn-zenbu.mjs --namae が 書いた・手で 並べて いません）',
    '# ★数★ ' + 台の名.length + '個'];
  fs.writeFileSync(名簿, 行.concat(台の名).join('\n') + '\n', 'utf-8');
  console.log('★名簿を 書いた … ' + 名簿 + '（' + 台の名.length + '個）★');
  console.log('★次に … pwsh -NoProfile -File docs/measured/toru-xlfn-zenbu.ps1★');
} else {
  if (!fs.existsSync(元)) {
    console.error('★先に `pwsh -NoProfile -File docs/measured/toru-xlfn-zenbu.ps1` を 走らせて ください★');
    process.exit(2);
  }
  /* ★zip を 解く★ */
  const b = fs.readFileSync(元);
  let eocd = -1;
  for (let i = b.length - 22; i >= 0 && i > b.length - 70000; i--) {
    if (b.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
  }
  const 数 = b.readUInt16LE(eocd + 10);
  let p = b.readUInt32LE(eocd + 16), xml = null;
  for (let k = 0; k < 数; k++) {
    if (b.readUInt32LE(p) !== 0x02014b50) break;
    const 法 = b.readUInt16LE(p + 10), 圧 = b.readUInt32LE(p + 20);
    const 名長 = b.readUInt16LE(p + 28), 追長 = b.readUInt16LE(p + 30), 注長 = b.readUInt16LE(p + 32);
    const 先 = b.readUInt32LE(p + 42);
    const 名 = b.toString('utf8', p + 46, p + 46 + 名長);
    if (名 === 'xl/worksheets/sheet1.xml') {
      const ln = b.readUInt16LE(先 + 26), le = b.readUInt16LE(先 + 28);
      const 中 = b.slice(先 + 30 + ln + le, 先 + 30 + ln + le + 圧);
      xml = (法 === 0 ? 中 : zlib.inflateRawSync(中)).toString('utf8');
      break;
    }
    p += 46 + 名長 + 追長 + 注長;
  }
  if (!xml) { console.error('★中の sheet1.xml が 読めない★'); process.exit(2); }

  /* ★実Excel が 書いた 式から 名前を 拾う★ */
  const 要る = new Set(), 要らない = new Set();
  const re = /<f[^>]*>([^<]*)<\/f>/g;
  let m;
  while ((m = re.exec(xml)) !== null) {
    const f = m[1];
    const x = /^_xlfn\.([A-Z0-9_.]+)\(/.exec(f);
    if (x) { 要る.add(x[1]); continue; }
    const y = /^([A-Z][A-Z0-9_.]*)\(/.exec(f);
    if (y) 要らない.add(y[1]);
  }

  /* ★今 うちが 持って いる 一覧★ */
  const io = fs.readFileSync(path.join(ROOT, 'lib/xlsx-io.js'), 'utf-8');
  const mm = /var XLFN = \[([\s\S]*?)\];/.exec(io);
  const 今 = new Set((mm ? mm[1] : '').match(/'([^']+)'/g).map((s) => s.slice(1, -1)));

  const 足りない = [...要る].filter((n) => !今.has(n)).sort();
  const 余分 = [...今].filter((n) => 要らない.has(n)).sort();

  console.log('');
  console.log('★★実Excel に 聞いた★★');
  console.log('  打てた 式 … ' + (要る.size + 要らない.size) + '個');
  console.log('  ★`_xlfn.` が 要る★ … ' + 要る.size + '個');
  console.log('  `_xlfn.` が 要らない … ' + 要らない.size + '個');
  console.log('');
  console.log('★今 うちの 一覧★ … ' + 今.size + '個');
  console.log('');
  console.log('★★★足りない（付け忘れ＝実Excel で #NAME? に なる）★★★ … ' + 足りない.length + '個');
  if (足りない.length) console.log('  ' + 足りない.join(' '));
  console.log('');
  console.log('★余分（付ける 要が 無いのに 付けて いる）★ … ' + 余分.length + '個');
  if (余分.length) console.log('  ' + 余分.join(' '));
  console.log('');
  /* ★直しに 使える 形で 出す★ */
  if (足りない.length) {
    console.log('★一覧に 足す 形★');
    console.log("  '" + 足りない.join("', '") + "'");
  }
}
