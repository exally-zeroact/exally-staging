/* mae-ato-kuraberu.mjs — ★「前」と「後」の 紙を 行ごとに 突き合わせる★（2026-09-08）
 *
 *  ★★指示役 2026-09-08 の 注文★★
 *    「★『判定が 変わった 行』と『値が 変わった 行』を ★両方★ 出して ください★」
 *    ★訳★ … #48（IM系）で 私は 78本、指示役は 83本 と 出した。
 *            差の 5本は ★値は 変わったが 判定は『違う』の まま★。
 *            今回は 害が 無かった（合った → 違う は 0本）が、
 *            ★次の 直しで「★合った のまま 値が 変わる★」と ★誰も 気づけない★★。
 *
 *  ★出す 数（★1つの 数に しない★）★
 *    ①★合っていたのに 壊れた 行★（＝一番 大事）
 *    ②直った 行／直った 関数
 *    ③★判定は 同じだが 値が 変わった 行★（＝今まで 見えていなかった 物）
 *    ④前だけ／後だけに 在る 行（＝紙が ずれていないか）
 *
 *  使い方: node mae-ato-kuraberu.mjs <前.tsv> <後.tsv>
 */
import fs from 'node:fs';

const 前道 = process.argv[2];
const 後道 = process.argv[3];
if (!前道 || !後道) {
  console.error('使い方: node mae-ato-kuraberu.mjs <前.tsv> <後.tsv>');
  process.exit(2);
}

const 読む = (道) => {
  const m = new Map();
  for (const l of fs.readFileSync(道, 'utf-8').split('\n')) {
    if (!l) continue;
    const p = l.split('\t');
    if (p.length < 2) continue;
    m.set(p[0], { 判定: p[1], 値: p.length >= 3 ? p[2] : '' });
  }
  return m;
};

const 前 = 読む(前道);
const 後 = 読む(後道);

const 壊れた = [], 直った = [], 値だけ = [], 前だけ = [], 後だけ = [];
for (const [式, a] of 前) {
  const b = 後.get(式);
  if (!b) { 前だけ.push(式); continue; }
  if (a.判定 === '★数えない★' || b.判定 === '★数えない★') continue;
  if (a.判定 !== b.判定) {
    if (a.判定 === '合った') 壊れた.push(式 + '   ' + a.判定 + '（' + a.値 + '） → ' + b.判定 + '（' + b.値 + '）');
    else if (b.判定 === '合った') 直った.push(式);
    else 壊れた.push('（合ってはいなかった）' + 式 + '   ' + a.判定 + ' → ' + b.判定);
  } else if (a.値 !== b.値) {
    値だけ.push(式 + '   ' + a.判定 + ' の まま   ' + a.値 + ' → ' + b.値);
  }
}
for (const 式 of 後.keys()) if (!前.has(式)) 後だけ.push(式);

const 関数 = (式) => 式.replace(/^=/, '').replace(/\(.*/, '');
const 個 = (a) => new Set(a.map(関数)).size;

console.log('★入れる 前と 後を 行ごとに 突き合わせた★');
console.log('  前 ' + 前.size + '行 ／ 後 ' + 後.size + '行');
console.log('');
console.log('  ★★①合っていたのに 壊れた 行 …… ' + 壊れた.length + '本★★');
console.log('  ②直った 行 …………………………… ' + 直った.length + '本（関数 ' + 個(直った) + '個）');
console.log('  ★③判定は 同じだが 値が 変わった 行 … ' + 値だけ.length + '本★');
console.log('  ④前だけに 在る 行 ' + 前だけ.length + '本 ／ 後だけに 在る 行 ' + 後だけ.length + '本');
console.log('');
if (壊れた.length) { console.log('★★①壊れた 行★★'); for (const x of 壊れた) console.log('    ' + x); console.log(''); }
if (直った.length) {
  console.log('★②直った 関数（道具の 出力そのまま）★');
  const c = {};
  for (const x of 直った) c[関数(x)] = (c[関数(x)] || 0) + 1;
  for (const k of Object.keys(c).sort((p, q) => c[q] - c[p])) console.log('    ' + k.padEnd(16) + c[k] + '本');
  console.log('');
}
if (値だけ.length) {
  console.log('★③判定は 同じだが 値が 変わった 行★（★これが 今まで 見えていなかった★）');
  for (const x of 値だけ.slice(0, 40)) console.log('    ' + x);
  if (値だけ.length > 40) console.log('    …（' + (値だけ.length - 40) + '本 略）');
  console.log('');
}
if (前だけ.length || 後だけ.length) {
  console.log('★④紙が ずれている★');
  for (const x of 前だけ.slice(0, 10)) console.log('    前だけ … ' + x);
  for (const x of 後だけ.slice(0, 10)) console.log('    後だけ … ' + x);
}
