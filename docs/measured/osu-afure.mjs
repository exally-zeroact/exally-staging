/* osu-afure.mjs — ★取った 紙と 自前の 溢れの 台を 突き合わせる★（2026-09-13）
 *
 *  ★★何の 為か★★
 *    `lib/shiki-afure.js`（溢れ＝1つの 式が 何マスにも 広がる）の 答えは
 *    全部 実Excel の 実測です。
 *    ★実Excel の 版が 変わった／測り方が 崩れた 時に ここが 赤に なります★
 *
 *  ★使い方★
 *    pwsh -NoProfile -File docs/measured/toru-afure.ps1   … ★実Excel に 打って 紙を 作る★
 *    node docs/measured/osu-afure.mjs                     … ★突き合わせる★
 *
 *  ★★この 台は 見張りでは ありません★★
 *    `tests/run.js` から は 回しません（★Excel が 要る＝この パソコンでしか 動かない★）。
 *    見張り（tests/shiki-afure.test.mjs）は ★実Excel 無しで★ 守ります。
 *
 *  ★★列の 名前で 読みます（番号で 読まない）★★
 *    2026-09-13 に 紙の 列が ★6列→8列★ に 増えました。★番号なら 黙って ずれて いました★。
 *
 *  ★★紙の 全部を 突き合わせては いません★★
 *    紙には `=SUM(A1:A3)`（溢れない）や `=ROW(A1:A3)`（まだ 作って いない 関数）も 在ります。
 *    ★この 台が 見るのは「溢れの 決まり」だけ★＝★下の 表に 書いた 物だけ★。
 *    ⇒★出来て いない 物を 出来た 顔で 混ぜない★
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/shiki-keisan.js'));
const A = require_(path.join(ROOT, 'lib/shiki-afure.js'));

const 紙 = path.join(ROOT, 'docs/measured/golden-afure-2026-09-13.tsv');
if (!fs.existsSync(紙)) {
  console.error('★先に これを 走らせて ください★');
  console.error('  pwsh -NoProfile -File docs/measured/toru-afure.ps1');
  process.exit(2);
}

/* ★材料★（実Excel に 置いたのと 同じ） A1=1 A2=2 A3=3 ／ B1=10 B2=20 B3=30 */
const A列 = A.溢れ([[K.数(1)], [K.数(2)], [K.数(3)]]);
const B列 = A.溢れ([[K.数(10)], [K.数(20)], [K.数(30)]]);
const AB = A.溢れ([[K.数(1), K.数(10)], [K.数(2), K.数(20)]]);

const 字に = (v) => {
  if (A.溢れか(v)) return v.並び.map((段) => 段.map(字に).join('|')).join(' / ');
  if (v.型 === '誤') return v.値;
  if (v.型 === '真偽') return v.値 ? 'TRUE' : 'FALSE';
  if (v.型 === '数') return K.数を字に(v.値);
  if (v.型 === '空') return '';
  return v.値;
};

/* ★紙の 式 → 自前で 出す 答え★（★溢れの 決まりを 見る 物だけ★） */
const 出す = {
  '=A1:A3': () => A列,
  '=A1:A3*2': () => A.重ねる(A列, K.数(2), (a, b) => K.つなぎ('*', a, b, {})),
  '=A1:B2': () => AB,
  '=A1:A3+B1:B3': () => A.重ねる(A列, B列, (a, b) => K.つなぎ('+', a, b, {})),
  '=A1:A3+B1:B2': () => A.重ねる(A列, A.溢れ([[K.数(10)], [K.数(20)]]), (a, b) => K.つなぎ('+', a, b, {})),
  '=A1:A3&"x"': () => A.重ねる(A列, K.字('x'), (a, b) => K.つなぎ('&', a, b, {})),
  '=-A1:A3': () => A.ひとつずつ(A列, (v) => K.前置き('-', v, {})),
  '=A1:A3=1': () => A.重ねる(A列, K.数(1), (a, b) => K.つなぎ('=', a, b, {})),
  '=@A1:A3': () => A.先頭(A列),
};

/* ★紙を 読む（列の 名前で 位置を 決める）★ */
const 行たち = fs.readFileSync(紙, 'utf-8').split(/\r?\n/);
const 見出し = 行たち.find((l) => l.startsWith('# 式'));
if (!見出し) { console.error('★紙に 見出しの 行が 無い★'); process.exit(2); }
const 列名 = 見出し.replace(/^#\s*/, '').split('\t');
const 位置 = (名) => {
  const i = 列名.indexOf(名);
  if (i < 0) { console.error('★紙に 列「' + 名 + '」が 無い★（' + 列名.join(' / ') + '）'); process.exit(2); }
  return i;
};
const c式 = 位置('式'), cジャマ = 位置('ジャマ'), cE1 = 位置('E1の字'), c先 = 位置('広がった先の字');

let 合 = 0;
const 違い = [], 見ない = [];
for (const l of 行たち) {
  if (!l || l.startsWith('#')) continue;
  const c = l.split('\t');
  if (c.length <= c先) continue;
  if (c[cジャマ] !== 'まっさら') continue;          /* ★邪魔の 話は 見張り側で 見る★ */
  const 式 = c[c式];
  if (!出す[式]) { 見ない.push(式); continue; }

  /* ★紙の「E1の字」と「広がった先の字」から 実Excel の 並びを 作る★ */
  const 並び = [c[cE1]];
  if (c[c先] !== '—') {
    for (const t of c[c先].split(' ')) {
      const m = /^([A-Z]+)(\d+)=(.*)$/.exec(t);
      if (m) 並び.push(m[3]);
    }
  }
  const 欲 = 並び.join(' / ');
  const 得 = 字に(出す[式]()).split(' / ').map((s) => s.split('|').join(' / ')).join(' / ');
  if (得 === 欲) 合++;
  else 違い.push(式 + ' … 紙 `' + 欲 + '` ／ ★うち `' + 得 + '`★');
}

console.log('');
console.log('★★締め★★ 見た ' + (合 + 違い.length) + '通り ／ 合った ' + 合 + ' ／ ★違う ' + 違い.length + '★');
console.log('　紙の 列 … ' + 列名.join(' / '));
console.log('　★わざと 見ない★ … ' + 見ない.length + '通り（' + 見ない.join(' ') + '）');
console.log('　　＝溢れない 式／まだ 作って いない 関数＝★出来て いない 物を 出来た 顔で 混ぜない★');
if (違い.length) {
  console.log('');
  console.log('★合わない 物の 実物★');
  違い.forEach((x) => console.log('  ' + x));
}
process.exit(違い.length ? 1 : 0);
