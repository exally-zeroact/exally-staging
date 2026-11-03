/* oddf-gyakusan.mjs — ★合わない 7本で ★どの 部品が いくつ なら 合うか★を 逆に 解く★（2026-09-18）
 *
 *  ★★なぜ★★
 *    ・形を ★9つ★ 試して ★全部 外れ★（今の台 33 が 一番 良い）
 *    ・残差は ★1回の 利払いの 0.03%〜2.4%★＝★端の かけら の 大きさ★
 *    ⇒★★当てずっぽうで 形を 選ぶのを やめます★★
 *    ⇒★★「どの 部品が いくつ なら 合うか」を 逆に 解きます★★
 *
 *  ★★組み立ての 式★★（`初回端数を組む`）
 *    価格 = 償還 / 割^((N-1)+DSC)
 *         + 券 * DFC / 割^DSC
 *         + Σ(k=2..N) 券 / 割^((k-1)+DSC)
 *         - 券 * A
 *    ⇒★DFC も A も ★1次★★＝★他の 2つを 正しいと 置けば 解けます★
 *    ⇒★DSC は 1次では ありません★＝★数えて 寄せます★
 *
 *  ★★これは 答えを 当てる 道具では ありません★★
 *    ・★「その 部品だけが 違う」と 決めて いません★
 *    ・★★3通り 全部 出して ★どれが きれいか★を 見るだけ★★
 *    ・★きれいとは ★1日ぶん★ や ★ちょうど 0.5★ の ような 形★
 *
 *  ★★見て いない 事★★
 *    ・★2つ 以上 同時に 違う 時は ここでは 出ません★
 *    ・★#NUM! の 行は 扱いません★
 *
 *  ★★★出た 物（2026-09-18・★一番 大きい 手掛かり★）★★★
 *    ★★ちょうど 整数の 日数★★が 出ました
 *      組B basis 2 ... ★DFC が ちょうど 1日 多い★（要 -1.0000日）
 *      組B basis 3 ... ★DFC が ちょうど 1日 多い★（要 -1.0000日）
 *      組C basis 0 ... ★DSC が ちょうど 1日 多い★（要 -1.0000日）
 *      組C basis 4 ... ★DSC が ちょうど 2日 多い★（要 -2.0000日）
 *      組B basis 0 ... 0.05日（★ほぼ 合って いる★）
 *    ⇒★★「形」の 話では ありません／★日を 1つ 多く 数えて います★★★
 *    ⇒★★どこかの 準利払日が 1日 ずれて います★★ の 見込み
 *       ＝組B の 準日 ... 今 2007-12-31（★初回 06-30 が 月末なので 月末に 揃えた★）
 *         ⇒★2007-12-30 なら 41日 -> 40日★＝★ちょうど 合います★
 *    ★但し 組C は それでは 説明が 付きません★（★2日 ずれる★）
 *    ⇒★★まだ 決まって いません★★＝★次は 準日の 作り方を 1つずつ★
 *
 *  ★★★組C（残り 3本）を 日数まで 下ろしました（2026-09-18）★★★
 *    ★組C★ 決済 2009-03-10 ／ 初回 2009-08-31 ／ f=4 ／ 満期 2015-02-28
 *    ★準利払日★ 2008-11-30 ／ 2009-02-28 ／ 2009-05-31 ／ 2009-08-31
 *      （★月末を 揃えても 揃えなくても 同じ★＝組B の 直しが 効きません）
 *    ★DSC の 中身★ ... 端（2009-03-10 -> 2009-05-31）÷ 90 ＋ まるごと 1 ＋ 手当て 1
 *      basis 0 ... 今 81日（30/360 US・D1=10 なので D2=31 の まま）
 *                  ★要 80日★
 *      basis 4 ... 今 80日（30E/360・D2=31 -> 30）
 *                  ★要 78日★
 *    ⇒★★basis ごとに 要る 日数が 違います★★
 *      ・終わりを 05-30 に すると basis 0 は 合う（80）が basis 4 が 合わない（80 != 78）
 *      ・終わりを 05-28 に すると basis 4 は 合う（78）が basis 0 が 合わない（78 != 80）
 *    ⇒★★日付を 動かしても 両方は 合いません＝★数え方★の 話です★★
 *    ⇒★★ここから 先は 当てません★★（★次は COUPDAYSNC / COUPDAYS で 出す 形を 試す★）
 *
 *  使い方: node docs/measured/kansuu46/oddf-gyakusan.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '../../..');
const require_ = createRequire(path.join(ROOT, 'package.json'));
const K = require_(path.join(ROOT, 'lib/formula-kane.js'));

const 紙道 = path.join(ROOT, 'docs/measured/kansuu46/golden-kane-2026-09-07.tsv');
const 行 = fs.readFileSync(紙道, 'utf-8').split(/\r?\n/).filter((l) => l && !l.startsWith('#'));

function 割る(s) {
  const 出 = []; let 深 = 0, 今 = '';
  for (const ch of s) {
    if (ch === '(') 深++;
    if (ch === ')') 深--;
    if (ch === ',' && 深 === 0) { 出.push(今); 今 = ''; continue; }
    今 += ch;
  }
  出.push(今);
  return 出;
}
const 数に = (s) => {
  const m = /^DATE\((\d+),(\d+),(\d+)\)/.exec(s.trim());
  return m ? K.日から数(+m[1], +m[2], +m[3]) : Number(s.trim());
};
const 日 = (n) => K.数から日(n);
const 字 = (n) => { const d = 日(n); return d.y + '-' + String(d.m).padStart(2, '0') + '-' + String(d.d).padStart(2, '0'); };

console.log('');
console.log('★★合わない 本で ★どの 部品が いくつ なら 合うか★★★');
console.log('');

let 本 = 0;
for (const l of 行) {
  const c = l.split('\t');
  if (c[0] !== 'ODDFPRICE') continue;
  const v = 割る(c[1].slice(c[1].indexOf('(') + 1, c[1].lastIndexOf(')'))).map(数に);
  const 正 = Number(c[2]);
  if (!isFinite(正)) continue;
  const b = v.length > 8 ? v[8] : 0, f = v[7];
  const 出 = K.初回端数の価格(v[0], v[1], v[2], v[3], v[4], v[5], v[6], v[7], b);
  if (typeof 出 !== 'number') continue;
  if (Math.abs(出 - 正) <= Math.max(1e-9, Math.abs(正) * 1e-9)) continue;
  本++;

  /* ★今の 部品を 外から 取る★（`初回端数の中身` が 出して います） */
  const 中 = K.初回端数の中身(v[0], v[1], v[2], v[3], v[7], b);
  const DFC = 中 && 中.DFC, A = 中 && 中.A, DSC = 中 && 中.DSC, N = 中 && 中.N;
  if (!(typeof DFC === 'number')) {
    console.log('  ★中身が 取れません★ ' + c[1].slice(0, 70) + '  ' + JSON.stringify(中));
    continue;
  }
  const 割 = 1 + v[5] / f, 券 = 100 * v[4] / f;

  /* ★DFC だけ 動かす★（1次） */
  const 要DFC = DFC + (正 - 出) * Math.pow(割, DSC) / 券;
  /* ★A だけ 動かす★（1次・符号は マイナス） */
  const 要A = A - (正 - 出) / 券;
  /* ★DSC だけ 動かす★（数えて 寄せる） */
  const 価 = (dsc) => {
    let p = v[6] / Math.pow(割, (N - 1) + dsc) + 券 * DFC / Math.pow(割, dsc);
    for (let k = 2; k <= N; k++) p += 券 / Math.pow(割, (k - 1) + dsc);
    return p - 券 * A;
  };
  let 下 = DSC - 2, 上 = DSC + 2;
  for (let i = 0; i < 200; i++) {
    const m = (下 + 上) / 2;
    if (価(m) > 正) 下 = m; else 上 = m;        /* 価は dsc について 単調 減少 */
  }
  const 要DSC = (下 + 上) / 2;

  /* ★1期の 長さ★（端の 期・`比()` の B1 と 同じ） */
  const 準 = null;
  const 長 = (b === 1) ? null : ((b === 3) ? 365 / f : 360 / f);

  console.log('  ★' + c[1].slice(11, 74) + '★');
  console.log('      basis ' + b + ' ／ f ' + f + ' ／ 発行 ' + 字(v[2]) + ' ／ 初回 ' + 字(v[3])
    + ' ／ 決済 ' + 字(v[0]));
  console.log('      ★差★ ' + (出 - 正).toExponential(4));
  console.log('      DFC  今 ' + DFC.toFixed(10) + '  要 ' + 要DFC.toFixed(10)
    + '  ★差 ' + (要DFC - DFC).toFixed(8) + '★'
    + (長 ? '（日で ' + ((要DFC - DFC) * 長).toFixed(4) + '）' : ''));
  console.log('      A    今 ' + A.toFixed(10) + '  要 ' + 要A.toFixed(10)
    + '  ★差 ' + (要A - A).toFixed(8) + '★'
    + (長 ? '（日で ' + ((要A - A) * 長).toFixed(4) + '）' : ''));
  console.log('      DSC  今 ' + DSC.toFixed(10) + '  要 ' + 要DSC.toFixed(10)
    + '  ★差 ' + (要DSC - DSC).toFixed(8) + '★'
    + (長 ? '（日で ' + ((要DSC - DSC) * 長).toFixed(4) + '）' : ''));
}
console.log('');
console.log('  ★合わない 本 ... ' + 本 + '本★');
console.log('');
console.log('★見る 所★');
console.log('  ・★差が ★ちょうど 1日★ や ★ちょうど 0.5★ に なって いる 部品が 在るか★');
console.log('  ・★同じ 部品が 7本とも 同じ 向きに ずれて いるか★');
console.log('');
console.log('★言えない 事★');
console.log('  ・★「その 部品だけが 違う」とは 決めて いません★');
console.log('  ・★2つ 以上 同時に 違う 時は ここでは 出ません★');
