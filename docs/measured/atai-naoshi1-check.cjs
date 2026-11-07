const path=require('path');const W=process.argv[2];
require(path.join(W,'lib/shiki-tsunagi.js'));require(path.join(W,'lib/shiki-kansuu.js'));
const H=require(path.join(W,'lib/shiki-hyou.js'));
const cases=[
 ['=SUMIF(A1:A5,">2",B1:C5)','120'],['=SUMIF(A1:A5,">2",B1)','120'],['=SUMIF(A1:A5,">2",C1:C2)','1200'],
 ['=SUMIF(A1:B5,">2",C1:C5)','1200'],['=AVERAGEIF(A1:A5,">2",B1:C5)','40'],['=SUMIF(A:B,">2",C:C)','1212'],
 ['=SUMIF(A1:A5,">2",B1:B5)','120'],['=SUMIF(A1:A5,">2")','12'],
];
let ng=0;
for(const [f,want] of cases){const t=new H.表();for(let r=1;r<=5;r++){t.打つ('A'+r,r);t.打つ('B'+r,r*10);t.打つ('C'+r,r*100);}
 t.打つ('BZ1',f);const got=String(t.字('BZ1'));if(got!==want)ng++;console.log(got===want?'ok':'NG',f,'台',got,'Excel',want);}
const e=[['=COUNTIF(A1:A7,"")','2'],['=COUNTIF(A1:A7,"=")','2'],['=COUNTIF(A1:A7,"<>3")','6'],['=COUNTIF(A1:A7,"<>")','5'],['=COUNTIF(A1:A7,Z1)','0'],
 ['=COUNTIFS(A1:A7,"")','2'],['=SUMIF(A1:A7,"",B1:B7)','0'],['=SUMIF(A1:A7,"<>3",B1:B7)','24'],['=AVERAGEIF(A1:A7,"<>3")','3'],
 ['=XLOOKUP(Z1,A1:A7,B1:B7)','0'],['=XLOOKUP(Z1,A1:A7,B1:B7,"無")','0'],['=MATCH(Z1,A1:A7,0)','#N/A'],['=VLOOKUP(Z1,A1:B7,2,FALSE)','#N/A'],['=XMATCH(Z1,A1:A7)','6'],
 ['=AVERAGEIFS(B1:B7,A1:A7,"<>3")','6']];
for(const [f,want] of e){const t=new H.表();for(let r=1;r<=5;r++){t.打つ('A'+r,r);t.打つ('B'+r,r*2);}
 t.打つ('BZ1',f);const got=String(t.字('BZ1'));if(got!==want)ng++;console.log(got===want?'ok':'NG',f,'台',got,'Excel',want);}
console.log('NG',ng);
