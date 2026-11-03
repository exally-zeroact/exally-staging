/* shirabe-oddl.mjs — ★ODDLPRICE の DC・A の 数え方を 実測から 決める★（2026-09-07・調べ物） */
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';

const ここ = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(ここ, '..', '..', '..');
const K = createRequire(import.meta.url)(path.join(ROOT, 'lib/formula-kane.js'));
const 末日 = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();

function 並びを作る(最p, 満p, f) {
  const 月 = 12 / f, 末 = 最p.d === 末日(最p.y, 最p.m), 並 = [最p];
  for (let k = 1; k < 500; k++) {
    const 通 = 最p.y * 12 + (最p.m - 1) + k * 月;
    const y2 = Math.floor(通 / 12), m2 = (通 % 12) + 1, last = 末日(y2, m2);
    const p = { y: y2, m: m2, d: 末 ? last : Math.min(最p.d, last) };
    並.push(p);
    if (K.日から数(p.y, p.m, p.d) >= K.日から数(満p.y, 満p.m, 満p.d)) break;
  }
  return 並;
}

const 金 = fs.readFileSync(path.join(ここ, 'golden-kane-2026-09-07.tsv'), 'utf-8')
  .split('\n').filter((l) => l.startsWith('ODDLPRICE')).map((l) => l.split('\t'));

for (const [, 式, 正] of 金) {
  if (!isFinite(Number(正))) continue;
  const 日 = [...式.matchAll(/DATE\((\d+),(\d+),(\d+)\)/g)].map((x) => K.日から数(+x[1], +x[2], +x[3]));
  const 残 = 式.replace(/DATE\(\d+,\d+,\d+\)/g, 'D').slice('=ODDLPRICE('.length, -1)
    .split(',').filter((x) => x !== 'D').map(Number);
  const [r, y, R, f] = 残;
  const b = (残.length > 4) ? 残[4] : 0;
  const 決p = K.数から日(日[0]), 満p = K.数から日(日[1]), 最p = K.数から日(日[2]);
  const 並 = 並びを作る(最p, 満p, f);
  const 券 = 100 * r / f;

  function 比(始, 終, 方式) {
    const s = K.日から数(始.y, 始.m, 始.d), e = K.日から数(終.y, 終.m, 終.d);
    let 和 = 0;
    for (let i = 0; i < 並.length - 1; i++) {
      const a = 並[i], c = 並[i + 1];
      const a2 = K.日から数(a.y, a.m, a.d), c2 = K.日から数(c.y, c.m, c.d);
      const s2 = Math.max(a2, s), e2 = Math.min(c2, e);
      if (e2 <= s2) continue;
      let 長;
      if (方式 === 'A') 長 = (b === 0 || b === 4) ? 360 / f : (c2 - a2);
      else if (方式 === 'B') 長 = (b === 1) ? (c2 - a2) : ((b === 3) ? 365 / f : 360 / f);
      else 長 = (b === 0 || b === 4) ? 360 / f : (K.日から数(並[1].y, 並[1].m, 並[1].d) - K.日から数(並[0].y, 並[0].m, 並[0].d));
      和 += K.日数(K.数から日(s2), K.数から日(e2), b) / 長;
    }
    return 和;
  }
  const 出 = [];
  for (const 方式 of ['A', 'B', 'C']) {
    const DC = 比(最p, 満p, 方式), A = 比(最p, 決p, 方式);
    const v = ((R + 券 * DC) / (1 + (DC - A) * y / f)) - 券 * A;
    if (Math.abs(v - Number(正)) < 1e-8) 出.push(方式);
  }
  /* ★A だけ「満期までの 割合 − 決済からの 割合」に する 案★ */
  const DCa = 比(最p, 満p, 'A');
  const DSCa = 比(決p, 満p, 'A');
  const v2 = ((R + 券 * DCa) / (1 + DSCa * y / f)) - 券 * (DCa - DSCa);
  if (Math.abs(v2 - Number(正)) < 1e-8) 出.push('DSCを直に');
  console.log('b=' + b + ' 期数=' + (並.length - 1) + '  正=' + Number(正).toFixed(8)
    + '  合った … ' + (出.length ? 出.join(' / ') : '★無し★')
    + '  (A方式=' + (((R + 券 * DCa) / (1 + (DCa - 比(最p, 決p, 'A')) * y / f)) - 券 * 比(最p, 決p, 'A')).toFixed(8)
    + ' / DSC直=' + v2.toFixed(8) + ')');
}
