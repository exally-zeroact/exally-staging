/* tsukuru-ookii-hon.mjs — ★大きい 本を 作る（測る 為の 材料）★ 2026-09-22
 *
 *  ★★なぜ 作るか★★
 *    ★司さんの 実物で 測れませんでした★
 *      ＝実物は `OneDrive` に 在り、★写し取りが 止められました★（Sensitive-Source）。
 *      ⇒★「測って いない」と 言ったままには しません★＝★同じ 大きさの 本を 自分で 作ります★。
 *    ★これは 実物の 代わりに なりません★
 *      ＝★式の 形も 繋がり方も 私が 決めた 物★です。
 *      ⇒★「実物で 測った」とは 書きません★。★ここで 分かるのは 「大きさで どう 効くか」だけ★。
 *
 *  ★★何を 大きくするか★★
 *    ・板の 枚数 ... ★10枚★（「板が 10枚 ある 本」も まだ 測って いませんでした）
 *    ・式の マス ... 1枚あたり ★2,000★ ＝ 合計 ★20,000★
 *    ・関数 ....... `SUM` / `IF` / `ROUND` / `VLOOKUP` を 混ぜる（★種類も 数える 対象★）
 *    ・板を またぐ 式 ... ★入れます★（つられて 変わる 所を 作る）
 *
 *  ★repo に 置きません★＝出来た 本は `%TEMP%`（★材料は 作り直せる★・道具だけ 残す）
 *
 *  走らせ方: node docs/measured/tsukuru-ookii-hon.mjs [板の数] [1枚の式の数]
 */
import path from 'node:path'; import fs from 'node:fs'; import os from 'node:os';
import crypto from 'node:crypto';
import { createRequire } from 'node:module';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../..');
const require_ = createRequire(pathToFileURL(path.join(ROOT, 'package.json')));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));

const 板の数 = Number(process.argv[2] || 10);
const 式の数 = Number(process.argv[3] || 2000);

const wb = XLSX.utils.book_new();
for (let i = 0; i < 板の数; i++) {
  const ws = {};
  let 最大行 = 0;
  /* ★元に なる 数★（式が 指す 先） */
  for (let r = 0; r < 200; r++) {
    ws[XLSX.utils.encode_cell({ r: r, c: 0 })] = { t: 'n', v: (r + 1) * (i + 1) };
    最大行 = Math.max(最大行, r);
  }
  /* ★式★（4種類を 混ぜる） */
  for (let k = 0; k < 式の数; k++) {
    const r = 200 + k;
    const 指す = (k % 200) + 1;
    let f;
    if (k % 4 === 0) f = 'SUM(A1:A' + 指す + ')';
    else if (k % 4 === 1) f = 'IF(A' + 指す + '>10,A' + 指す + '*2,0)';
    else if (k % 4 === 2) f = 'ROUND(A' + 指す + '/3,2)';
    else f = 'VLOOKUP(A' + 指す + ',A1:A200,1,FALSE)';
    ws[XLSX.utils.encode_cell({ r: r, c: 1 })] = { t: 'n', f: f, v: 0 };
    最大行 = Math.max(最大行, r);
  }
  /* ★板を またぐ 式★＝★1マス 直すと 別の 板が 動く★ を 作る */
  if (i > 0) {
    ws[XLSX.utils.encode_cell({ r: 0, c: 3 })] = { t: 'n', f: "SUM('板1'!A1:A200)", v: 0 };
    ws[XLSX.utils.encode_cell({ r: 1, c: 3 })] = { t: 'n', f: "'板1'!A1*2", v: 0 };
  }
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: 最大行, c: 3 } });
  XLSX.utils.book_append_sheet(wb, ws, '板' + (i + 1));
}

/* ★★置き場は 引数で 変えられます★★（2026-09-22）
     `os.tmpdir()` へ 書けない 事が 在ります（★書ける 所は 走らせ方で 変わります★）。
     ⇒★書けない 時に 「作れません」で 終わらせない★＝置き場を 渡せるように します。 */
const 置き場 = process.argv[4] || os.tmpdir();
const 出先 = path.join(置き場, 'exally-ookii-' + 板の数 + 'mai-' + 式の数 + 'shiki.xlsx');
const 中 = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
fs.writeFileSync(出先, 中);
console.log('★作りました★ ' + 出先);
console.log('  板 ' + 板の数 + '枚 ／ 式 ' + (板の数 * 式の数 + (板の数 - 1) * 2).toLocaleString() + 'マス'
  + ' ／ ' + 中.length.toLocaleString() + ' バイト');
console.log('  sha256 ' + crypto.createHash('sha256').update(中).digest('hex'));
console.log('  ★実物の 代わりでは ありません★＝式の 形は 私が 決めた 物です。');
