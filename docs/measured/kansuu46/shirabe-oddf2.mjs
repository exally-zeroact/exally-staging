/* shirabe-oddf2.mjs — ★合わない 組の「割り引く 回数」を 逆算する★（2026-09-07・調べ物）
 *  ⇒ 決めつけずに ★実Excel の 答えから x を 逆に 求めて★、x に 意味が 有るかを 見る
 */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const K = createRequire(import.meta.url)(path.join(ROOT, 'lib/formula-kane.js'));

const 末日 = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
function 並びを作る(初p, 発p, f) {
  const 月 = 12 / f, 末 = 初p.d === 末日(初p.y, 初p.m), 並 = [初p];
  for (let k = -1; k > -500; k--) {
    const 通 = 初p.y * 12 + (初p.m - 1) + k * 月;
    const y2 = Math.floor(通 / 12), m2 = (通 % 12) + 1, last = 末日(y2, m2);
    const p = { y: y2, m: m2, d: 末 ? last : Math.min(初p.d, last) };
    並.unshift(p);
    if (K.日から数(p.y, p.m, p.d) <= K.日から数(発p.y, 発p.m, 発p.d)) break;
  }
  return 並;
}

const 金 = fs.readFileSync(path.join(ここ, 'golden-kane-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l.startsWith('ODDFPRICE')).map((l) => l.split('\t'));

for (const [, 式, 正] of 金) {
  if (!isFinite(Number(正))) continue;
  const 日 = [...式.matchAll(/DATE\((\d+),(\d+),(\d+)\)/g)].map((x) => K.日から数(+x[1], +x[2], +x[3]));
  const 残 = 式.replace(/DATE\(\d+,\d+,\d+\)/g, 'D').slice('=ODDFPRICE('.length, -1)
    .split(',').filter((x) => x !== 'D').map(Number);
  const [r, y, R, f] = 残;
  const b = (残.length > 4) ? 残[4] : 0;
  const 決p = K.数から日(日[0]), 満p = K.数から日(日[1]), 発p = K.数から日(日[2]), 初p = K.数から日(日[3]);
  if (K.日から数(発p.y, 発p.m, 発p.d) >= K.日から数(決p.y, 決p.m, 決p.d)) continue;
  const 並 = 並びを作る(初p, 発p, f);
  const NC = 並.length - 1;
  const 比 = (始, 終) => {
    const s = K.日から数(始.y, 始.m, 始.d), e = K.日から数(終.y, 終.m, 終.d);
    let 和 = 0;
    for (let i = 0; i < 並.length - 1; i++) {
      const a = 並[i], c = 並[i + 1];
      const a2 = K.日から数(a.y, a.m, a.d), c2 = K.日から数(c.y, c.m, c.d);
      const s2 = Math.max(a2, s), e2 = Math.min(c2, e);
      if (e2 <= s2) continue;
      const 長 = (b === 1) ? (c2 - a2) : ((b === 3) ? 365 / f : 360 / f);
      和 += K.日数(K.数から日(s2), K.数から日(e2), b) / 長;
    }
    return 和;
  };
  const DFC = 比(発p, 初p), A = 比(発p, 決p), 素 = DFC - A;
  const N = 1 + K.利払回数(初p, 満p, f);
  const 割 = 1 + y / f, 券 = 100 * r / f;
  const 価 = (x) => {
    let 出 = R / Math.pow(割, (N - 1) + x);
    出 += 券 * DFC / Math.pow(割, x);
    for (let k = 2; k <= N; k++) 出 += 券 / Math.pow(割, (k - 1) + x);
    return 出 - 券 * A;
  };
  let lo = -1, hi = 6;
  for (let i = 0; i < 200; i++) { const m = (lo + hi) / 2; if (価(m) > Number(正)) lo = m; else hi = m; }
  const x = (lo + hi) / 2;
  const 発が準の日か = 並.some((p) => K.日から数(p.y, p.m, p.d) === K.日から数(発p.y, 発p.m, 発p.d));
  console.log('b=' + b + ' NC=' + NC + ' 発が準日=' + (発が準の日か ? 'はい' : 'いいえ')
    + '  素の x=' + 素.toFixed(6) + '  逆算 x=' + x.toFixed(6) + '  差=' + (x - 素).toFixed(6));
}
