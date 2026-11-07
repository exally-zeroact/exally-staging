/* 写し字（* ? ~）の 確かめ＝utsushiji-excel-2026-11-07.tsv（実Excel 16.0 build 20527）と 台を 比べる
   使い方 `node docs/measured/utsushiji-check.cjs <作業木>` */
const path=require('path'),fs=require('fs');const W=path.resolve(process.argv[2]||process.cwd());
require(path.join(W,'lib/shiki-tsunagi.js'));require(path.join(W,'lib/shiki-kansuu.js'));
const H=require(path.join(W,'lib/shiki-hyou.js'));
const rows=fs.readFileSync(path.join(__dirname,'utsushiji-excel-2026-11-07.tsv'),'utf8').split(/\r?\n/).filter(l=>l&&l[0]!=='#').map(l=>l.split('\t'));
const 盤=['apple','Apple','ab','a*b',5,null,'a?c',15];
let ng=0;
for(const [f,want] of rows){const t=new H.表();盤.forEach((v,i)=>{if(v!==null)t.打つ('A'+(i+1),v);t.打つ('B'+(i+1),i+1);});
 t.打つ('BZ1',f);const got=String(t.字('BZ1'));const ok=got===want||(isFinite(+got)&&Math.abs(+got-+want)<1e-9);if(!ok)ng++;console.log(ok?'ok':'NG',f,'台',got,'Excel',want);}
console.log('NG',ng,'/',rows.length);process.exitCode=ng?1:0;
