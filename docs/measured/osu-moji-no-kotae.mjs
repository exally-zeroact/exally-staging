/* osu-moji-no-kotae.mjs — ★文字を 返す 式が 数に 化けないか★（2026-09-11）
 *
 *  ★★何が 起きて いたか★★
 *    `=TEXT(A1,"0.00")` の 答えは 実Excel では ★文字の "3.00"★（LEN=4／ISTEXT=TRUE）。
 *    うちも ★答えは 文字で 持って いました★が、
 *    `_hfGetDisplay` が `String(val)` に した 時点で ★型が 消え★、
 *    画面が それを ★数として 読み直して★ ★"3"★ と 出して いました。
 *
 *  ★★TEXT だけでは ありません★★
 *    LEFT／RIGHT／MID／TRIM／SUBSTITUTE／CONCATENATE／REPT／`&`
 *    ⇒★文字を 返す 式は 全部★（実Excel で 11本 確かめた・全部 String）
 *    （`="007"` だけは ★頭の ゼロ★の 決まりで たまたま 守られて いた）
 *
 *  ★★この 道具の 数は 画面の 数です★★
 *    本物の ブラウザで ★マスを 押して 字を 打ち★、
 *    ★画面が 本当に 描いた 字★（fillText）を 取って 実Excel の 紙と 突き合わせます。
 *    ⇒★作り直した 字では ありません★
 *
 *  ★この 道具は エンジンを 建てません★＝計算するのは 画面の 中の 本番の エンジン。
 *  ★司さんの 実物には 触りません★＝手元に 立てた 配信の book.html だけ。
 *
 *  使い方:
 *    ① pwsh -NoProfile -File docs/measured/toru-moji-no-kotae.ps1  … 実Excel に 聞く
 *    ② node docs/measured/osu-moji-no-kotae.mjs                    … 押して 突き合わせる
 */
import http from 'node:http'; import fs from 'node:fs'; import path from 'node:path';
import { borrow, launch, unmeasured } from '../../scripts/_borrow-playwright.mjs';
import { fileURLToPath } from 'node:url';
const ROOT=path.resolve(path.join(path.dirname(fileURLToPath(import.meta.url)), '..', '..'));
const 紙=path.join(ROOT,'docs/measured/golden-moji-no-kotae-2026-09-11.tsv');
const 実=[];
for(const l of fs.readFileSync(紙,'utf-8').split(/\r?\n/)){ if(!l||l.startsWith('#'))continue; const c=l.split('\t'); if(c.length<5)continue; 実.push({式:c[0],字:c[1],型:c[2],何:c[4]}); }
const 型={'.html':'text/html; charset=utf-8','.js':'text/javascript; charset=utf-8','.css':'text/css; charset=utf-8','.json':'application/json; charset=utf-8','.png':'image/png','.svg':'image/svg+xml'};
const s=http.createServer((rq,rs)=>{const d=decodeURIComponent(String(rq.url).split('?')[0]);const f=path.join(ROOT,d.replace(/^\/+/,''));if(!f.startsWith(ROOT)||!fs.existsSync(f)||fs.statSync(f).isDirectory()){rs.statusCode=404;return rs.end('no');}rs.setHeader('content-type',型[path.extname(f).toLowerCase()]||'application/octet-stream');fs.createReadStream(f).pipe(rs);});
await new Promise(r=>s.listen(0,'127.0.0.1',r));
const c=await borrow('moji-no-kotae','chromium');
if(!c){ unmeasured('moji-no-kotae','chromium'); process.exit(0); }
const b=await launch('moji-no-kotae',c,{},'chromium');
if(!b){ unmeasured('moji-no-kotae','chromium'); process.exit(0); }
const p=await b.newPage({viewport:{width:1200,height:760}});
await p.goto('http://127.0.0.1:'+s.address().port+'/book.html',{waitUntil:'load'});
await p.evaluate(()=>{document.body.classList.remove('exally-locked');const o=document.getElementById('loginOv');if(o)o.style.display='none';});
await p.waitForTimeout(600);
const 居所=(r,c)=>p.evaluate(([r,c])=>{const cv=document.getElementById('grid-canvas');const b=cv.getBoundingClientRect();return {x:b.left+colX(c)+20,y:b.top+rowY(r)+8};},[r,c]);
const 打つ=async(r,c,v)=>{const q=await 居所(r,c);await p.mouse.click(q.x,q.y);await p.keyboard.type(v);await p.keyboard.press('Enter');await p.keyboard.press('Escape');await p.waitForTimeout(110);};
await 打つ(0,0,'3');
for(let i=0;i<実.length;i++) await 打つ(i+2,1,実[i].式);
await p.evaluate(()=>{window.__描いた=[];const 元=CanvasRenderingContext2D.prototype.fillText;CanvasRenderingContext2D.prototype.fillText=function(t,x,y){try{window.__描いた.push([String(t),x,y]);}catch(e){}return 元.apply(this,arguments);};render();});
await p.waitForTimeout(300);
const 出=await p.evaluate((n)=>{const 結=[];for(let i=0;i<n;i++){const r=i+2,c=1;const x=colX(c),y=rowY(r),w=cW(c),h=rH(r);let 字='';for(const [t,tx,ty] of (window.__描いた||[])) if(tx>=x-1&&tx<=x+w+1&&ty>=y-1&&ty<=y+h+1){字=t;break;}結.push(字);}return 結;},実.length);
let 合=0,違=0;
console.log('');
for(let i=0;i<実.length;i++){
  const よい = 出[i]===実[i].字;
  if(よい) 合++; else 違++;
  console.log((よい?'  ok  ':'  NG  ')+実[i].式.padEnd(30)+'実Excel="'+実[i].字+'" うち="'+出[i]+'"');
}
console.log('\n★★締め★★ 全 '+(合+違)+'本 ／ 合った '+合+' ／ ★違う '+違+'★');
await b.close(); s.close(); process.exit(違?1:0);
