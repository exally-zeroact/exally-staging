/* ★見立て「差 ＝ はみ出しの 日数（その basis）÷ 1期の 長さ」を 字だけで 試す★
   ★実Excel は 叩きません★（紙は 6枠ぶん 手元に 在ります） */
import { createRequire } from 'node:module';
const require_ = createRequire('C:/Users/zeroa/exally-prod/package.json');
const K = require_('C:/Users/zeroa/exally-prod/lib/formula-kane.js');
const D = (y,m,d) => ({y,m,d});
const 数 = (p) => K.日から数(p.y,p.m,p.d);
const 末日 = (y,m) => new Date(Date.UTC(y,m,0)).getUTCDate();
const 月足す = (p,k) => { const 通=p.y*12+(p.m-1)+k, y=Math.floor(通/12), m=(通%12)+1;
  return {y,m,d:Math.min(p.d,末日(y,m))}; };
/* ★貼り付く★（前の 日から 進める） */
const 並びS = (最,満,f) => { const st=12/f; const q=[最]; let c=最;
  while (数(c) < 数(満)) { c = 月足す(c, st); q.push(c); } return q; };
/* ★月末そろえ★（今の lib） */
const 並びM = (最,満,f) => { const st=12/f, 末=最.d===末日(最.y,最.m); const q=[最]; let k=0;
  while (true) { k++; const t=月足す({y:最.y,m:最.m,d:1}, k*st);
    const p={y:t.y,m:t.m,d:末?末日(t.y,t.m):Math.min(最.d,末日(t.y,t.m))};
    q.push(p); if (数(p) >= 数(満)) break; } return q; };

const 組 = [
 {名:'組1',f:4,最:D(2008,11,30),満:D(2009,8,31), 決:D(2009,3,10), 差:{0:1/45,1:3/92,2:3/92,3:3/92,4:1/45}},
 {名:'組2',f:2,最:D(2007,10,15),満:D(2008,6,15), 決:D(2008,2,7),  差:{0:0,1:0,2:0,3:0,4:0}},
 {名:'組3',f:2,最:D(2008,11,30),満:D(2009,5,31), 決:D(2009,1,15), 差:{0:0,1:1/181,2:1/181,3:1/181,4:0}},
 {名:'組4',f:4,最:D(2008,11,30),満:D(2009,7,15), 決:D(2009,1,15), 差:{0:1/45,1:0,2:0,3:0,4:0}},
 {名:'組5',f:1,最:D(2009,1,31), 満:D(2010,2,28), 決:D(2009,6,15), 差:{0:1/360,1:0,2:0,3:0,4:0}},
 {名:'組6',f:2,最:D(2008,11,30),満:D(2010,5,31), 決:D(2009,1,15), 差:{0:0,1:1/181,2:1/181,3:1/181,4:0}},
];
const 実 = (a,b) => 数(b)-数(a);
for (const 貼 of [true,false]) {
  console.log('\n===== 準利払日 ＝ ' + (貼?'★貼り付く★':'★月末そろえ（今の lib）★') + ' =====');
  for (const g of 組) {
    const q = 貼 ? 並びS(g.最,g.満,g.f) : 並びM(g.最,g.満,g.f);
    let i最後 = 0; for (let i=0;i<q.length;i++) if (数(q[i]) <= 数(g.満)) i最後 = i;
    const 準 = q[i最後];
    console.log('  '+g.名+' 並 '+q.map(p=>p.y+'-'+p.m+'-'+p.d).join(' ')
      +'  最後の準 '+準.y+'-'+準.m+'-'+準.d);
    for (const b of [0,1,2,3,4]) {
      const 実日 = 実(準,g.満);
      const 基日 = K.日数(準, g.満, b);              /* ★その basis で 数え直す★ */
      const 長A = (b===0||b===4) ? 360/g.f : (i最後+1<q.length ? 実(q[i最後],q[i最後+1]) : NaN);
      const 出 = 基日 / 長A;
      const 期 = g.差[b];
      const ok = Math.abs(出-期) < 1e-9 || (期===0 && 基日===0);
      console.log('     b'+b+' はみ 実'+実日+'日 / basis '+基日+'日  長'+長A
        +'  ⇒ '+出.toFixed(9)+'  実測 '+期.toFixed(9)+'  '+(ok?'★合う★':'★★違う★★'));
    }
  }
}

/* ★★分母を 出す★★（★30行 とも 数える＝分母を 出さない 緑は 嘘★） */
console.log('\n★★数えます（貼り付く 並び）★★');
let 合=0, 全=0, 差在=0, 差在合=0;
for (const g of 組) {
  const q = 並びS(g.最,g.満,g.f);
  let i最後=0; for (let i=0;i<q.length;i++) if (数(q[i])<=数(g.満)) i最後=i;
  for (const b of [0,1,2,3,4]) {
    const 基日 = K.日数(q[i最後], g.満, b);
    const 長A = (b===0||b===4) ? 360/g.f : (i最後+1<q.length ? 実(q[i最後],q[i最後+1]) : NaN);
    const 出 = 基日/長A, 期 = g.差[b];
    const ok = Math.abs(出-期) < 1e-9;
    全++; if (ok) 合++;
    if (期 !== 0) { 差在++; if (ok) 差在合++; }
  }
}
console.log('  ★見立てが 合う … ' + 合 + ' / ' + 全 + '★（差が 在る 行だけなら ' + 差在合 + ' / ' + 差在 + '）');

/* ★★別の 見立て★★ … 満期と 最後の 準利払日が ★同じ 月★か */
console.log('\n★★別の 見立て「満期と 最後の 準利払日が 同じ 月か」★★（★差が 出るか 出ないか だけ★）');
let c=0;
for (const g of 組) {
  const q = 並びS(g.最,g.満,g.f);
  let i最後=0; for (let i=0;i<q.length;i++) if (数(q[i])<=数(g.満)) i最後=i;
  const 同月 = (q[i最後].y===g.満.y && q[i最後].m===g.満.m);
  const 差在り = [0,1,2,3,4].some(b=>g.差[b]!==0);
  const ok = (同月===差在り); if (ok) c++;
  console.log('  '+g.名+' 最後の準 '+q[i最後].y+'-'+q[i最後].m+'-'+q[i最後].d
    +' ／ 満 '+g.満.y+'-'+g.満.m+'-'+g.満.d
    +'  同じ月='+(同月?'はい':'いいえ')+' ／ 差が 出た='+(差在り?'はい':'いいえ')+'  '+(ok?'★合う★':'★違う★'));
}
console.log('  ★' + c + ' / 6 組★');
