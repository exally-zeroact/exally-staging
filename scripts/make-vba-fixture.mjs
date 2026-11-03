/* make-vba-fixture.mjs — ★試験用の .xlsm を うちで作る★（客の実物は 使わない）
 * =============================================================================
 * ★決まり（指示役 2026-08-28）★
 *   ・★試験用の .xlsm を1本 作って repo に入れる★（★客の実物は 使わない★）
 *   ・うちに .xlsm は 0本だったので ★中身も うちで書く★
 *   ・VBAは ★動かさない★。ここで作るのは ★読む練習のための 紙★。
 *
 * 作る物 … xl/vbaProject.bin（CFB）の中に
 *   VBA/dir（目次・MS-OVBA圧縮）／VBA/Module1・Module2・Sheet1（中身・MS-OVBA圧縮）
 *   ＋ Excelが要る最低限の流れ（PROJECT / VBA/_VBA_PROJECT）
 *
 * 使い方: node scripts/make-vba-fixture.mjs
 *   → tests/fixtures/vba-sample.xlsm を 作り直す
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const Vba = require_(path.join(ROOT, 'lib/vba.js'));

const CFB = XLSX.CFB;

/* ★中身は うちで書いた物★（毎月の締め作業を まねた 3本）
   ★中身の元は tests/fixtures/vba-src/*.bas★＝★cp932（Shift_JIS）で 置いてある★。
   VBAの中身は Unicode では 置かれていない（日本語Excelは 932）。
   ここで 1バイトずつ 作ると 日本語が 化けた（実測 2026-08-28）ので、
   ★元のバイトを そのまま 読んで そのまま 圧縮する★。 */
const 元の所 = path.join(ROOT, 'tests/fixtures/vba-src');
const モジュール = [
  { 名: 'Module1', 種類: 0x0021, 生: fs.readFileSync(path.join(元の所, 'Module1.bas')) },
  { 名: 'Module2', 種類: 0x0021, 生: fs.readFileSync(path.join(元の所, 'Module2.bas')) },
  { 名: 'Module3', 種類: 0x0021, 生: fs.readFileSync(path.join(元の所, 'Module3.bas')) },
  { 名: 'Sheet1', 種類: 0x0022, 生: fs.readFileSync(path.join(元の所, 'Sheet1.bas')) },
];

/* ── dir（目次）を 組み立てる ── */
function 記録(id, 中) {
  const b = [id & 0xFF, (id >> 8) & 0xFF, 中.length & 0xFF, (中.length >> 8) & 0xFF,
    (中.length >> 16) & 0xFF, (中.length >> 24) & 0xFF];
  return b.concat(Array.from(中));
}
const 字 = (s) => Array.from(s).map((c) => c.charCodeAt(0) & 0xFF);
const 字16 = (s) => Array.from(s).flatMap((c) => [c.charCodeAt(0) & 0xFF, (c.charCodeAt(0) >> 8) & 0xFF]);
const 数4 = (n) => [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF];

/** ★マクロの入れ物（vbaProject.bin）を 組み立てる★
 *  ★試験からも 呼ぶ★＝同じ物を 2か所に 書かない（写すと ずれる）
 *  @param {{字の種類?:number, 詰めない?:boolean}} 決め
 *    字の種類 … PROJECTCODEPAGE（既定 932＝日本語）。読めない値を 入れて 試すためにある
 *    詰めない … true なら ★縮めずに 入れる★（＝中身は同じ・詰め方だけ 違う物を 作る）
 */
export function 入れ物を組み立てる(決め) {
  決め = 決め || {};
  const 字の種類 = 決め.字の種類 === undefined ? 932 : 決め.字の種類;
  const 詰めない = !!決め.詰めない;
  /* ★目次の名前を わざと ずらす★＝「目次の名前」と「中身の1行目」が 食い違う物を作る。
     ★読み違えた時に 起きる事★を その場で 作って、うちが 気づけるかを 測るため。 */
  const 名前をずらす = !!決め.名前をずらす;
  /* ★目次の 途中を 読めなくする★＝行き止まりに 気づけるかを 測るため */
  const 目次を壊す = !!決め.目次を壊す;

  let dir = [];
  dir = dir.concat(記録(0x0001, 数4(0x0004)));                 /* PROJECTSYSKIND */
  dir = dir.concat(記録(0x0003, [字の種類 & 0xFF, (字の種類 >> 8) & 0xFF]));  /* PROJECTCODEPAGE */
  dir = dir.concat(記録(0x0004, 字('ExallyTest')));            /* PROJECTNAME */
  dir = dir.concat(記録(0x000F, [モジュール.length & 0xFF, 0x00]));  /* PROJECTMODULES */
  /* ★PROJECTVERSION(0x0009)★ ここだけ 大きさが 嘘（「4」と書いて 中身は 6バイト）。
     ★実物には 必ず 在る★ので 見本にも 入れる＝入れないと この穴を 一度も 測れない
     （実物5本 ぜんぶ ここで 目次が 総崩れになり「マクロ0本」と 出ていた。2026-08-28）。 */
  dir = dir.concat([0x09, 0x00, 0x04, 0x00, 0x00, 0x00,
    0x01, 0x00, 0x00, 0x00, 0x00, 0x00]);
  if (目次を壊す) {
    /* 大きさが とんでもない 記録を 1つ 差し込む＝ここから先は 読めない */
    dir = dir.concat([0x99, 0x00, 0xFF, 0xFF, 0xFF, 0x7F]);
  }
  dir = dir.concat(記録(0x0013, [0xFF, 0xFF]));                /* PROJECTCOOKIE */
  for (const m of モジュール) {
    dir = dir.concat(記録(0x0019, 字(名前をずらす ? (m.名 + 'X') : m.名)));  /* MODULENAME */
    dir = dir.concat(記録(0x0047, 字16(m.名)));                /* MODULENAME_Unicode */
    dir = dir.concat(記録(0x001A, 字(m.名)));                  /* MODULESTREAMNAME */
    dir = dir.concat(記録(0x0032, 字16(m.名)));                /* …_Unicode */
    dir = dir.concat(記録(0x0031, 数4(0)));                    /* MODULEOFFSET＝0（頭から） */
    dir = dir.concat(記録(0x001E, 数4(0)));                    /* MODULEHELPCONTEXT */
    dir = dir.concat(記録(m.種類, []));                        /* MODULETYPE */
    dir = dir.concat(記録(0x002B, []));                        /* MODULEEND */
  }
  dir = dir.concat([0x10, 0x00, 0x00, 0x00, 0x00, 0x00]);      /* Terminator */

  const dir圧縮 = Vba.圧縮(new Uint8Array(dir));

  /* ── CFB を 組み立てる ── */
  const cfb = CFB.utils.cfb_new({ root: 'VBA' });
  CFB.utils.cfb_add(cfb, '/VBA/dir', Array.from(dir圧縮));
  for (const m of モジュール) {
    const z = 詰めない ? 生のまま入れる(new Uint8Array(m.生)) : Vba.圧縮(new Uint8Array(m.生));
    CFB.utils.cfb_add(cfb, '/VBA/' + m.名, Array.from(z));
  }
  CFB.utils.cfb_add(cfb, '/VBA/_VBA_PROJECT', [0xCC, 0x61, 0x00, 0x00, 0x00, 0x00]);
  CFB.utils.cfb_add(cfb, '/PROJECT', 字([
    'ID="{00000000-0000-0000-0000-000000000000}"',
    'Module=Module1', 'Module=Module2', 'Module=Module3', 'Document=Sheet1/&H00000000',
    'Name="ExallyTest"', '',
  ].join('\r\n')));
  /* ★vbaraw は バイト列で 渡す★ */
  const bin = Buffer.from(CFB.write(cfb, { type: 'array' }));

  return Buffer.from(CFB.write(cfb, { type: 'array' }));
}

/** ★縮めずに 入れる★＝MS-OVBAの「そのまま」の かたまり（頭の 15ビット目を 立てない） */
function 生のまま入れる(元) {
  const 出 = [0x01];
  for (let i = 0; i < 元.length; i += 4096) {
    const 中 = 元.subarray(i, Math.min(i + 4096, 元.length));
    const 頭 = ((中.length - 1) & 0x0FFF) | 0x3000;          /* 縮めた印(0x8000)を 立てない */
    出.push(頭 & 0xFF, (頭 >> 8) & 0xFF);
    for (let k = 0; k < 中.length; k++) 出.push(中[k]);
  }
  return Uint8Array.from(出);
}

/* ★直に叩いた時だけ 書き出す★
   （試験が この道具を import する。読み込んだだけで 見本が 書き換わると
     ★試験が 自分で作った物を 測る★事になり、置いてある見本を 一度も 見なくなる） */
if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  /* ── .xlsm に 入れる ── */
  const bin = 入れ物を組み立てる();
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([['日付', '数量', '単価', '金額'], [46023, 2, 500, 1000]]);
  XLSX.utils.book_append_sheet(wb, ws, '明細');
  wb.vbaraw = bin;
  const 出す = path.join(ROOT, 'tests/fixtures/vba-sample.xlsm');
  /* ★書き出すのは 自分で★（writeFile は この作りだと fs を 持っていない） */
  const バイト = XLSX.write(wb, { bookType: 'xlsm', bookVBA: true, type: 'buffer' });
  fs.writeFileSync(出す, バイト);

  const 大きさ = fs.statSync(出す).size;
  console.log('');
  console.log('[試験用の .xlsm を 作った]（★うちで書いた中身・客の実物は 使っていない★）');
  console.log('  置いた所 … tests/fixtures/vba-sample.xlsm（' + 大きさ + 'バイト）');
  console.log('  マクロ … ' + モジュール.map((m) => m.名 + '(' + m.生.toString('binary').split('\r\n').length + '行)').join(' / '));
  console.log('');
}
