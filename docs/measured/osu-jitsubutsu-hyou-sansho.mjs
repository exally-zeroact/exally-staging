/* osu-jitsubutsu-hyou-sansho.mjs — ★表の 参照が どれだけ A1 に 直るか 数える★（2026-09-15）
 *
 *  ★★なぜ★★（指示役1 2026-09-15）
 *    司さんの 実物の 式 15,799本の うち
 *      ★表の 参照（Table1[…]）… 12,218本（77.3%）★
 *      ★板を またぐ（シート名!）… 3,286本（20.8%）★
 *    ⇒★関数が 8個 揃っても この 1冊は 動きません★
 *    ⇒★でも 直す 道は 既に 在るかも しれない★＝`lib/table-refs.js`
 *      （2026-08-18 に 作って 在る＝★作る前に repo を 探せ★）
 *    ⇒★何本 直るのか★を 数えます。★ここで ⑥の 大きさが 決まります★。
 *
 *  ★★読むだけ★★
 *    `fs.readFileSync` で 読むだけ（★書きません★）／★前後で 大きさと 更新時刻を 確かめます★
 *    ★出すのは 数と 形だけ★＝★会社名・金額・人の 名前は 1文字も 出しません★
 *      （表の 名前・列の 名前も ★客が 付けた 字★なので ★伏せて 数だけ★ 出します）
 *
 *  使い方: node docs/measured/osu-jitsubutsu-hyou-sansho.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const XLSX = require_(path.join(ROOT, 'lib/xlsx.full.min.js'));
const TR = require_(path.join(ROOT, 'lib/table-refs.js'));
const ZS = require_(path.join(ROOT, 'lib/zip-surgeon.js'));

const 本 = 'C:\\Users\\zeroa\\kyukyu-0907\\代行計算表2026.xlsb';
if (!fs.existsSync(本)) { console.error('★本体が 無い★'); process.exit(2); }
const 前 = fs.statSync(本);
const bytes = new Uint8Array(fs.readFileSync(本));
const wb = XLSX.read(bytes, { type: 'array', cellFormula: true });

/* ★表の 参照か★／★板を またぐか★（★引用符の 中は 見ない★） */
const 裸に = (f) => String(f).replace(/"(?:[^"]|"")*"/g, '""');
const 表か = (f) => /\[[^\]]*\]/.test(裸に(f));
/* ★板を またぐか★（★`#REF!` は 板では ない★＝2026-09-15）
     ★前は ここだけ 直って いませんでした★＝★作る道が 2本★
     ⇒★`tests/hakaridai-mon.test.mjs` が 2本目を 赤に します★ */
const 板か = (f) => /[A-Za-z0-9_\u3000-\u9fff']+!/.test(裸に(f).replace(/#REF!/g, ''));

function 数える(取り) {
  let 全 = 0, 表 = 0, 板 = 0, 表も板も = 0;
  const 形 = {};
  for (const 名 of wb.SheetNames) {
    const ws = wb.Sheets[名];
    if (!ws || !ws['!ref']) continue;
    const R = XLSX.utils.decode_range(ws['!ref']);
    for (let r = R.s.r; r <= R.e.r; r++) {
      for (let c = R.s.c; c <= R.e.c; c++) {
        const s = ws[XLSX.utils.encode_cell({ r, c })];
        if (!s || !s.f) continue;
        const f = 取り(名, r, c, s.f);
        if (f === null) continue;
        全++;
        const t = 表か(f), b = 板か(f);
        if (t) 表++;
        if (b) 板++;
        if (t && b) 表も板も++;
        if (t) {
          /* ★形だけ 数える★＝★中の 名前は 出しません★（客が 付けた 字） */
          for (const m of 裸に(f).match(/[A-Za-z0-9_.\u3000-\u9fff]*\[[^\]]*\]/g) || []) {
            const 形字 = m
              .replace(/^[A-Za-z0-9_.\u3000-\u9fff]+/, '表名')
              .replace(/\[#([A-Za-z]+)\]/g, '[#$1]')          /* #Data 等は 決まり字＝残す */
              .replace(/\[@?[^#\]]+\]/g, (x) => (x.charAt(1) === '@' ? '[@列名]' : '[列名]'));
            形[形字] = (形[形字] || 0) + 1;
          }
        }
      }
    }
  }
  return { 全, 表, 板, 表も板も, 形 };
}

const 前の数 = 数える((n, r, c, f) => '=' + f);

console.log('# ★表の 参照が どれだけ A1 に 直るか★（2026-09-15）');
console.log('#   ★読むだけ★／★出すのは 数と 形だけ★（表の 名前・列の 名前も 伏せて 在ります）');
console.log('');
console.log('★直す 前（SheetJS が 出した まま）★');
console.log('  式 ' + 前の数.全 + '本 ／ ★表の 参照 ' + 前の数.表 + '本★ ／ 板またぎ ' + 前の数.板
  + '本 ／ 両方 ' + 前の数.表も板も + '本');
console.log('  ★形ごと（多い順・★名前は 伏せて 在ります★）★');
Object.keys(前の数.形).sort((a, b) => 前の数.形[b] - 前の数.形[a]).slice(0, 15)
  .forEach((k) => console.log('    ' + String(前の数.形[k]).padStart(6) + '本  ' + k));

console.log('');
console.log('★`lib/table-refs.js` に 直させます★（2026-08-18 に 作って 在る 道具）');
const 結 = await TR.resolve(bytes, 'xlsb', wb, ZS);
const 直し = (結 && 結.fixes) ? 結.fixes : (結 || {});
const 直った数 = Object.keys(直し).length;
console.log('  ★直した マス … ' + 直った数 + '★');
if (結 && 結.断り) {
  console.log('  ★断った 訳★');
  const d = 結.断り;
  if (typeof d === 'object') {
    Object.keys(d).forEach((k) => console.log('    ' + k + ' … ' + JSON.stringify(d[k]).slice(0, 120)));
  }
}

const 後の数 = 数える((n, r, c, f) => {
  const k = n + '|' + r + ',' + c;
  return 直し[k] !== undefined ? String(直し[k]) : ('=' + f);
});
console.log('');
console.log('★直した 後★');
console.log('  式 ' + 後の数.全 + '本 ／ ★表の 参照が 残る ' + 後の数.表 + '本★ ／ 板またぎ ' + 後の数.板 + '本');
console.log('  ⇒★A1 に なった … ' + (前の数.表 - 後の数.表) + '本★（'
  + (前の数.表 ? Math.round((前の数.表 - 後の数.表) / 前の数.表 * 1000) / 10 : 0) + '%）');
if (後の数.表) {
  console.log('  ★残った 形（★名前は 伏せて 在ります★）★');
  Object.keys(後の数.形).sort((a, b) => 後の数.形[b] - 後の数.形[a]).slice(0, 15)
    .forEach((k) => console.log('    ' + String(後の数.形[k]).padStart(6) + '本  ' + k));
}

const 後 = fs.statSync(本);
console.log('');
console.log('★本体★ … ' + 後.size + ' バイト … '
  + ((前.mtimeMs === 後.mtimeMs && 前.size === 後.size) ? '★動いて いません★' : '★★動いた★★'));
