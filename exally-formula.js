/**
 * exally-formula.js - Exally 数式計算エンジン（完全版・A案）
 * ================================================================
 * 【役割】Excel 500+ 関数の独自実装（96関数全部入り）
 * 【切り出し元】book.html line 1006〜1596（約590行）
 * 【バージョン】2026-04-20 完全版
 * 【設計】
 *   - hf (HyperFormula) は内部変数 _hf として保持
 *   - book.html 側で hf 初期化後に initExallyFormula(hf) を1回呼ぶだけ
 *   - 呼び出し側の関数シグネチャは一切変更なし
 *   - Node.js / ブラウザ両対応（末尾 module.exports）
 * 【使用箇所】
 *   - book.html        (メインブック画面)
 *   - kyuuryoumeisai.html (給料明細)
 *   - seikyusyo.html   (請求書)
 *   - mitsumoriyo.html (見積書)
 * ================================================================
 */

// ================================================================
// 【HyperFormula インスタンスホルダー】
// book.html 側で hf を作成後、initExallyFormula(hf) で渡す
// ================================================================
var _hf = null;

function initExallyFormula(hfInstance) {
  _hf = hfInstance;
}

function _hfSid(sheet){
  if(typeof sheet==='number') return sheet;
  try{ var id=_hf.getSheetId(sheet); return (id!==null&&id!==undefined)?id:0; }catch(e){return 0;}
}
function addSheetToEngine(name) {
  if(!_hf) return;
  try { _hf.addSheet(name); } catch(e) {}
}

// ================================================================
// 内部ヘルパー
// ================================================================
function _hfGetDisplay(sheet, r, c) {
  try {
    var val = _hf.getCellValue({sheet:_hfSid(sheet), row:r, col:c});
    if(val===null||val===undefined) return '';
    if(typeof val==='object'&&val.type) return HF_ERR[val.type]||('#'+val.type);
    return String(val);
  } catch(e) { return '#ERR'; }
}
function _toRC(addr) {
  var m = addr.match(/^([A-Z]+)(\d+)$/i);
  if(!m) return null;
  var col=0;
  for(var i=0;i<m[1].length;i++) col=col*26+(m[1].toUpperCase().charCodeAt(i)-64);
  return {r:parseInt(m[2])-1, c:col-1};
}
function _getRangeVals(sheet, rangeStr) {
  var m = rangeStr.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
  if(!m||!_hf) return [];
  var s=_toRC(m[1]),e=_toRC(m[2]),vals=[];
  for(var r=s.r;r<=e.r;r++)
    for(var c=s.c;c<=e.c;c++){
      var v=_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:c});
      if(typeof v==='number') vals.push(v);
    }
  return vals;
}
function _getRangeAll(sheet, rangeStr) {
  var m = rangeStr.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);
  if(!m||!_hf) return [];
  var s=_toRC(m[1]),e=_toRC(m[2]),vals=[];
  for(var r=s.r;r<=e.r;r++)
    for(var c=s.c;c<=e.c;c++)
      vals.push(_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:c}));
  return vals;
}
function _getSingleVal(sheet, ref) {
  var rc=_toRC(ref);
  if(!rc||!_hf) return null;
  return _hf.getCellValue({sheet:_hfSid(sheet),row:rc.r,col:rc.c});
}

// ================================================================
// 引数の式を値にする(★入れ子の関数呼び出しに対応)
//   以前は引数を素朴なカンマ分割で切っていたため、TEXTJOIN(",",TRUE,SORT(E1:E6,1,-1)) のような
//   入れ子があると『1,-1)』という無意味な文字列を黙って返していた。
//   ・引数の切り分けは _parseFuncArgs(括弧と引用符を数える)を使う
//   ・値にできない式は HyperFormula の calculateFormula に計算させる
//     (getCellValue は配列の先頭1個しか返さないが、calculateFormula は配列のまま返す)
//   ・★エラーはエラーのまま返す。壊れた文字列を返さない。
// ================================================================
var _ERRTXT = {DIV_BY_ZERO:'#DIV/0!',NUM:'#NUM!',NA:'#N/A',VALUE:'#VALUE!',REF:'#REF!',NAME:'#NAME?',CYCLE:'#CYCLE!',NULL:'#NULL!',SPILL:'#SPILL!'};
function _errText(v){
  if(typeof HF_ERR!=='undefined' && HF_ERR[v.type]) return HF_ERR[v.type];
  return _ERRTXT[v.type] || ('#'+v.type);
}
function _isErrObj(v){ return !!v && typeof v==='object' && !Array.isArray(v) && !!v.type; }
function _isErrBox(x){ return !!x && typeof x==='object' && !Array.isArray(x) && typeof x.err==='string'; }

function _argList(sheet, arg) {
  arg = String(arg).trim();
  var q = arg.match(/^"([\s\S]*)"$/);
  if(q) return [q[1].replace(/""/g,'"')];
  if(/^[A-Z]+\d+:[A-Z]+\d+$/i.test(arg)) return _getRangeAll(sheet, arg);
  if(/^[A-Z]+\d+$/i.test(arg))            return [_getSingleVal(sheet, arg)];
  if(/^-?\d+(\.\d+)?$/.test(arg))         return [parseFloat(arg)];
  if(/^(TRUE|FALSE)$/i.test(arg))         return [arg.toUpperCase()==='TRUE'];
  if(!_hf || typeof _hf.calculateFormula!=='function') return {err:'#VALUE!'};
  var v;
  try { v = _hf.calculateFormula(arg.charAt(0)==='=' ? arg : '='+arg, _hfSid(sheet)); }
  catch(e){ return {err:'#VALUE!'}; }
  if(_isErrObj(v)) return {err:_errText(v)};
  if(Array.isArray(v)){
    var out=[];
    for(var i=0;i<v.length;i++){
      var row=Array.isArray(v[i]) ? v[i] : [v[i]];
      for(var j=0;j<row.length;j++){
        if(_isErrObj(row[j])) return {err:_errText(row[j])};
        out.push(row[j]);
      }
    }
    return out;
  }
  return [v];
}
function _argScalar(sheet, arg) {
  var l = _argList(sheet, arg);
  if(_isErrBox(l)) return l;
  return l.length ? l[0] : null;
}
// 「その式が NAME(...) だけで出来ているか」を括弧を数えて確かめ、中身の引数文字列を返す。
//  ★ /^INT\s*\((.+)\)$/ のような貪欲な正規表現だと INT(NOW())-TODAY() の
//    最後の ")" まで飲み込んで別物になる(実際にこれで揮発性の確認を壊した)。
function _wholeCallArgs(f, name) {
  var m = new RegExp('^' + name + '\\s*\\(', 'i').exec(f);
  if(!m) return null;
  var i = m[0].length, depth = 1, ins = false;
  for(; i<f.length; i++){
    var c = f.charAt(i);
    if(c==='"'){ ins=!ins; continue; }
    if(ins) continue;
    if(c==='(') depth++;
    else if(c===')'){ depth--; if(depth===0) break; }
  }
  if(depth!==0) return null;
  if(i !== f.length-1) return null;      // 閉じ括弧が式の末尾でない=この関数だけの式ではない
  return f.slice(m[0].length, i);
}
function _argNum(v) {
  if(typeof v==='number') return v;
  if(typeof v==='boolean') return v?1:0;
  if(typeof v==='string' && v.trim()!=='' && !isNaN(v)) return parseFloat(v);
  return null;
}

// ================================================================
// JS実装関数群
// ================================================================

// --- TEXT書式 ---------------------------------------------------
//  ★以前は #,##0 / 0.00 / 0% / 0.00% / ¥#,##0 の5通りしか見ておらず、
//    それ以外(0.0% や yyyy/mm/dd、正負で分ける #,##0;(#,##0))は書式を無視して
//    数値をそのまま返していた。0.1235 が「0.1%」になる(=100倍違う)状態だった。
//  ★対応できない書式では null を返す。呼び出し側はHyperFormulaに任せる。
//    黙って String(num) を返すと「独自層を足したせいで悪くなる」ことがあるため
//    (実際に日付書式でそうなっていた)。

// Excelの丸め=四捨五入(0から遠い方向)。JSのMath.roundは負数で挙動が違うので符号を分ける。
// 12.35 のような2進で表せない数は e記法の文字列経由で丸める(toFixedだと12.3になる)。
function _fmtRound(n, d) {
  var sign = n<0 ? -1 : 1, a = Math.abs(n);
  var r = +(Math.round(+(a + 'e' + d)) + 'e' + (-d));
  return sign * r;
}
// 「正の書式;負の書式」を引用符の外の ; で分ける
function _fmtSections(fmt) {
  var out=[], cur='', ins=false;
  for(var i=0;i<fmt.length;i++){
    var c=fmt.charAt(i);
    if(c==='"'){ ins=!ins; cur+=c; continue; }
    if(c===';'&&!ins){ out.push(cur); cur=''; continue; }
    cur+=c;
  }
  out.push(cur);
  return out;
}
function _fmtLit(s) { return String(s).replace(/"([^"]*)"/g,'$1').replace(/\\(.)/g,'$1'); }
function _fmtIsDate(fmt) {
  var f = String(fmt).replace(/"[^"]*"/g,'');
  return /[ymdhs]/i.test(f) && !/[#0]/.test(f);
}
// シリアル値 → 日付(1900系。Excelの1900年閏年バグ域=シリアル60以前は扱わない)
function _fmtDate(serial, fmt) {
  var days = Math.floor(serial);
  var ms   = Math.round((serial - days) * 86400) * 1000;
  var d    = new Date(Date.UTC(1899,11,30) + days*86400000 + ms);
  var Y=d.getUTCFullYear(), Mo=d.getUTCMonth()+1, D=d.getUTCDate();
  var H=d.getUTCHours(), Mi=d.getUTCMinutes(), S=d.getUTCSeconds();
  var p2=function(v){ return (v<10?'0':'')+v; };
  return String(fmt).replace(/"[^"]*"|yyyy|yy|mm|m|dd|d|hh|h|ss|s/gi, function(t){
    if(t.charAt(0)==='"') return t.slice(1,-1);
    switch(t.toLowerCase()){
      case 'yyyy': return String(Y);
      case 'yy':   return p2(Y%100);
      case 'mm':   return p2(Mo);
      case 'm':    return String(Mo);
      case 'dd':   return p2(D);
      case 'd':    return String(D);
      case 'hh':   return p2(H);
      case 'h':    return String(H);
      case 'ss':   return p2(S);
      case 's':    return String(S);
      default:     return t;
    }
  });
}
function _fmtNumber(num, fmt) {
  var secs = _fmtSections(fmt);
  var useNeg = num<0 && secs.length>1;
  var f = useNeg ? secs[1] : secs[0];
  var v = useNeg ? Math.abs(num) : num;
  var m = f.match(/^([^#0]*?)([#0][#0,]*(?:\.[#0]+)?)(%?)([^#0]*)$/);
  if(!m) return null;
  var pre=m[1], core=m[2], pct=m[3], suf=m[4];
  if(pct) v = v*100;
  var dot = core.indexOf('.');
  var dec = dot>=0 ? core.length-dot-1 : 0;
  var grp = core.indexOf(',')>=0;
  var intCore = (dot>=0 ? core.slice(0,dot) : core).replace(/,/g,'');
  var minInt = (intCore.match(/0/g)||[]).length;
  var r = _fmtRound(v, dec);
  var sign = r<0 ? '-' : '';
  var parts = Math.abs(r).toFixed(dec).split('.');
  while(parts[0].length < minInt) parts[0] = '0' + parts[0];
  if(grp) parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return sign + _fmtLit(pre) + parts.join('.') + pct + _fmtLit(suf);
}
function _applyTextFormat(num, fmt) {
  var bare = String(fmt).replace(/"[^"]*"/g,'');
  if(/m{3,}|d{3,}/i.test(bare)) return null;   // 月名/曜日名(mmm・dddd)は未対応=HFに任せる
  if(_fmtIsDate(fmt)) return _fmtDate(num, fmt);
  return _fmtNumber(num, fmt);
}

// --- 統計 ---
function _jsRank(val,vals,order){var s=vals.slice().sort(function(a,b){return order===0?b-a:a-b;});var r=s.indexOf(val)+1;return r>0?r:'#N/A';}
function _jsPercentile(vals,k){var s=vals.slice().sort(function(a,b){return a-b;}),n=s.length,idx=k*(n-1),lo=Math.floor(idx),hi=Math.ceil(idx);if(lo===hi)return s[lo];return s[lo]+(idx-lo)*(s[hi]-s[lo]);}
function _jsQuartile(vals,q){return _jsPercentile(vals,[0,0.25,0.5,0.75,1][q]);}
function _jsMode(vals){var freq={};vals.forEach(function(v){if(typeof v==='number')freq[v]=(freq[v]||0)+1;});var keys=Object.keys(freq);if(!keys.length)return '#N/A';var max=0,mode=null;keys.forEach(function(k){if(freq[k]>max){max=freq[k];mode=parseFloat(k);}});return max>1?mode:'#N/A';}
function _jsTrimmean(vals,pct){var s=vals.slice().sort(function(a,b){return a-b;});var trim=Math.floor(s.length*pct/2);var t=s.slice(trim,s.length-trim);return t.reduce(function(a,b){return a+b;},0)/t.length;}
function _jsPercentrank(vals,x,sig){var s=vals.slice().sort(function(a,b){return a-b;}),n=s.length;var below=s.filter(function(v){return v<x;}).length;var rank=below/(n-1);var d=Math.pow(10,sig||3);return Math.floor(rank*d)/d;}
function _jsFrequency(data,bins){var r=new Array(bins.length+1).fill(0);data.forEach(function(v){var p=false;for(var i=0;i<bins.length;i++){if(v<=bins[i]){r[i]++;p=true;break;}}if(!p)r[bins.length]++;});return r;}
function _jsKurt(vals){var n=vals.length;if(n<4)return '#DIV/0!';var mean=vals.reduce(function(a,b){return a+b;})/n;var s2=vals.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/(n-1);var s=Math.sqrt(s2);var sum4=vals.reduce(function(a,v){return a+Math.pow((v-mean)/s,4);},0);return (n*(n+1))/((n-1)*(n-2)*(n-3))*sum4-3*(n-1)*(n-1)/((n-2)*(n-3));}
function _jsPermut(n,k){var r=1;for(var i=n;i>n-k;i--)r*=i;return r;}
function _jsPermutationa(n,k){return Math.pow(n,k);}
function _jsProb(vals,probs,lower,upper){if(upper===undefined)upper=lower;var sum=0;vals.forEach(function(v,i){if(v>=lower&&v<=upper)sum+=probs[i];});return sum;}
function _jsBinomDistRange(n,p,s1,s2){function bp(t,pr,k){var c=1;for(var i=0;i<k;i++)c=c*(t-i)/(i+1);return c*Math.pow(pr,k)*Math.pow(1-pr,t-k);}s2=s2===undefined?s1:s2;var sum=0;for(var k=s1;k<=s2;k++)sum+=bp(n,p,k);return sum;}

// --- 回帰・予測 ---
function _jsSlope(ys,xs){var n=ys.length,sx=0,sy=0,sxx=0,sxy=0;for(var i=0;i<n;i++){sx+=xs[i];sy+=ys[i];sxx+=xs[i]*xs[i];sxy+=xs[i]*ys[i];}return(n*sxy-sx*sy)/(n*sxx-sx*sx);}
function _jsIntercept(ys,xs){var slope=_jsSlope(ys,xs),n=ys.length,sx=ys.reduce(function(a,_,i){return a+xs[i];},0)/n,sy=ys.reduce(function(a,b){return a+b;})/n;return sy-slope*sx;}
function _jsForecast(x,ys,xs){return _jsIntercept(ys,xs)+_jsSlope(ys,xs)*x;}
function _jsLinest(ys,xs){return[_jsSlope(ys,xs),_jsIntercept(ys,xs)];}
function _jsLogest(ys,xs){var lys=ys.map(function(v){return Math.log(v);});return[Math.exp(_jsSlope(lys,xs)),Math.exp(_jsIntercept(lys,xs))];}
function _jsTrend(ys,xs,newXs){var s=_jsSlope(ys,xs),i=_jsIntercept(ys,xs);return newXs.map(function(x){return i+s*x;});}
function _jsGrowth(ys,xs,newXs){var lg=_jsLogest(ys,xs);return newXs.map(function(x){return lg[1]*Math.pow(lg[0],x);});}

// --- 行列 ---
function _jsMdeterm(m){var n=m.length;if(n===1)return m[0][0];if(n===2)return m[0][0]*m[1][1]-m[0][1]*m[1][0];var det=0;for(var c=0;c<n;c++){var sub=m.slice(1).map(function(r){return r.filter(function(_,ci){return ci!==c;});});det+=Math.pow(-1,c)*m[0][c]*_jsMdeterm(sub);}return det;}
function _jsMinverse(m){var det=_jsMdeterm(m);if(Math.abs(det)<1e-10)return null;if(m.length===2)return[[m[1][1]/det,-m[0][1]/det],[-m[1][0]/det,m[0][0]/det]];return null;}

// --- 文字列 ---
function _jsConcat(vals){return vals.map(function(v){return v===null||v===undefined?'':String(v);}).join('');}
function _jsTextjoin(delim,ignoreEmpty,vals){var f=ignoreEmpty?vals.filter(function(v){return v!==null&&v!==undefined&&v!=='';}) :vals;return f.map(function(v){return String(v===null||v===undefined?'':v);}).join(delim);}
// Excelの丸め=0から遠い方へ(ROUND(-2.5,0)=-3)。JSの Math.round は -2 になるので使えない。
//  2.675*100=267.49999… のような二進小数の誤差を吸収してから丸める。
function _xlRound(x,d){
  var p=Math.pow(10,d), v=x*p, eps=Math.abs(v)*1e-12;
  var r = v>=0 ? Math.floor(v+0.5+eps) : Math.ceil(v-0.5-eps);
  return r/p;
}
// FIXED(数値,[桁数],[桁区切りを付けない])。★桁数は負も取る(=FIXED(1234.5,-2) → "1,200")。結果は文字列。
function _jsFixed(num,dec,noCommas){
  dec = (dec===undefined||dec===null) ? 2 : Math.trunc(dec);
  var n = _xlRound(num, dec);
  var str = n.toFixed(Math.max(0,dec));
  if(!noCommas){var p=str.split('.');p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,',');str=p.join('.');}
  return str;
}
function _jsDollar(num,dec){dec=dec===undefined?2:Math.max(0,dec);var abs=Math.abs(num),str=abs.toFixed(dec);var p=str.split('.');p[0]=p[0].replace(/\B(?=(\d{3})+(?!\d))/g,',');str='$'+p.join('.');return num<0?'('+str+')':str;}
function _jsYen(num,dec){dec=dec===undefined?0:Math.max(0,dec);var v=Math.abs(num);var str=v.toFixed(dec).replace(/\B(?=(\d{3})+(?!\d))/g,',');return'¥'+str;}
function _jsLenb(s){return s.split('').reduce(function(a,c){return a+(c.charCodeAt(0)>0x7F?2:1);},0);}
function _jsLeftb(s,b){var r='',bytes=0;for(var i=0;i<s.length&&bytes<b;i++){var w=s.charCodeAt(i)>0x7F;if(bytes+(w?2:1)>b)break;r+=s[i];bytes+=w?2:1;}return r;}
function _jsRightb(s,b){return _jsLeftb(s.split('').reverse().join(''),b).split('').reverse().join('');}
function _jsMidb(s,start,len){return _jsLeftb(s.substring(start-1),len);}
function _jsFindb(find,within,start){start=start||1;var pos=within.indexOf(find,start-1);return pos>=0?pos+1:'#VALUE!';}
function _jsReplaceb(text,start,numBytes,repl){return text.substring(0,start-1)+repl+text.substring(start-1+numBytes);}
// ── 全角/半角(ASC / DBCS) ────────────────────────────────────────────
//  ★半角→全角の関数の【本名は DBCS】。JIS は日本語UIの表示名で、
//    xlsx の中身/US-English構文では DBCS。JIS のまま書き出すと Excel で #NAME? になる(実測)。
//    日本語UIの JIS( は convertFormula で DBCS( に直す＝打つ側はどちらでもよい。
//  ★対応表は順番で対応させる。ズレたら黙って誤変換するので、読み込み時に長さを検査する。
var _KZ_BASE = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲンァィゥェォャュョッー。「」、・゛゜';
var _KH_BASE = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝｧｨｩｪｫｬｭｮｯｰ｡｢｣､･ﾞﾟ';
var _KZ_DAKU = 'ガギグゲゴザジズゼゾダヂヅデドバビブベボヴヷヺ';
var _KH_DAKU = 'ｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾊﾋﾌﾍﾎｳﾜｦ';
var _KZ_HAND = 'パピプペポ';
var _KH_HAND = 'ﾊﾋﾌﾍﾎ';
if(_KZ_BASE.length!==_KH_BASE.length || _KZ_DAKU.length!==_KH_DAKU.length || _KZ_HAND.length!==_KH_HAND.length){
  throw new Error('全角/半角カナの対応表がズレています(exally-formula.js)。直してから使うこと。');
}
// 全角→半角。濁点は分解する(ガ→ｶﾞ)。漢字・ひらがなは半角が無いので触らない。
function _jsAsc(s){
  var out='';
  for(var i=0;i<s.length;i++){
    var c=s.charAt(i), k;
    if((k=_KZ_DAKU.indexOf(c))>=0){ out+=_KH_DAKU.charAt(k)+'ﾞ'; continue; }
    if((k=_KZ_HAND.indexOf(c))>=0){ out+=_KH_HAND.charAt(k)+'ﾟ'; continue; }
    if((k=_KZ_BASE.indexOf(c))>=0){ out+=_KH_BASE.charAt(k); continue; }
    var cc=c.charCodeAt(0);
    if(cc>=0xFF01 && cc<=0xFF5E){ out+=String.fromCharCode(cc-0xFEE0); continue; }  // 全角英数記号
    if(cc===0x3000){ out+=' '; continue; }                                          // 全角スペース
    out+=c;
  }
  return out;
}
// 半角→全角。濁点は合成する(ｶﾞ→ガ)。
function _jsDbcs(s){
  var out='';
  for(var i=0;i<s.length;i++){
    var c=s.charAt(i), n=s.charAt(i+1), k;
    if(n==='ﾞ' && (k=_KH_DAKU.indexOf(c))>=0){ out+=_KZ_DAKU.charAt(k); i++; continue; }
    if(n==='ﾟ' && (k=_KH_HAND.indexOf(c))>=0){ out+=_KZ_HAND.charAt(k); i++; continue; }
    if((k=_KH_BASE.indexOf(c))>=0){ out+=_KZ_BASE.charAt(k); continue; }
    var cc=c.charCodeAt(0);
    if(cc>=0x21 && cc<=0x7E){ out+=String.fromCharCode(cc+0xFEE0); continue; }       // 半角英数記号
    if(cc===0x20){ out+='　'; continue; }                                        // 半角スペース
    out+=c;
  }
  return out;
}
var _jsJis = _jsDbcs;   // 旧名(日本語UI名)。中身は同じ物を指す＝二重実装にしない。
// TEXTBEFORE/TEXTAFTER(文字列, 区切り, [出現回数], [大小区別], [末尾一致], [見つからない時])
//  ★出現回数は負も取る(-1=後ろから1つ目)。見つからない時は既定 #N/A、第6引数があればそれを返す。
function _tbPositions(t,d,ci){
  var hay = ci ? t.toLowerCase() : t, ned = ci ? d.toLowerCase() : d;
  var out=[], i=0;
  while(ned!=='' && (i=hay.indexOf(ned,i))!==-1){ out.push(i); i++; }
  return out;
}
function _tbPick(t,d,inst,ci){
  inst = (inst===undefined||inst===null||inst===0) ? 1 : Math.trunc(inst);
  var pos=_tbPositions(String(t),String(d),ci);
  return inst>0 ? pos[inst-1] : pos[pos.length+inst];
}
function _jsTextbefore(text,delim,inst,matchMode,matchEnd,ifNotFound){
  var t=String(text), d=String(delim);
  var p=_tbPick(t,d,inst,!!matchMode);
  if(p===undefined) return ifNotFound===undefined ? '#N/A' : ifNotFound;
  return t.substring(0,p);
}
function _jsTextafter(text,delim,inst,matchMode,matchEnd,ifNotFound){
  var t=String(text), d=String(delim);
  var p=_tbPick(t,d,inst,!!matchMode);
  if(p===undefined) return ifNotFound===undefined ? '#N/A' : ifNotFound;
  return t.substring(p+d.length);
}
function _jsTextsplit(text,colDelim,rowDelim){if(rowDelim)return text.split(rowDelim).map(function(r){return r.split(colDelim);});return text.split(colDelim);}
function _jsValuetotext(v){if(typeof v==='string')return'"'+v+'"';if(typeof v==='boolean')return v?'TRUE':'FALSE';return String(v);}
function _jsArraytotext(vals){return'{'+vals.map(_jsValuetotext).join(',')+'}';}

// --- 日付 ---
// DATEVALUE。受ける表記は【実Excelに聞いた物だけ】(推測で広げない):
//   2026/7/31 ・ 2026-07-31 ・ 2026年7月31日  … いずれも 46234(実測)
function _jsDateValue(str){
  var s=String(str).trim();
  var m=s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/)
     || s.match(/^(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日$/);
  if(!m) return null;
  var y=parseInt(m[1],10), mo=parseInt(m[2],10), da=parseInt(m[3],10);
  if(mo<1||mo>12||da<1||da>31) return null;
  var d=new Date(y,mo-1,da);
  if(d.getFullYear()!==y || d.getMonth()!==mo-1 || d.getDate()!==da) return null; // 2026-02-30 等
  return Math.round((d-new Date(1899,11,30))/86400000);
}
// NUMBERVALUE(文字列,[小数点],[桁区切り])。区切りを引数で指定できるのが VALUE との違い。
function _jsNumbervalue(text,decSep,grpSep){
  var d=(decSep===undefined||decSep===null||decSep==='') ? '.' : String(decSep).charAt(0);
  var g=(grpSep===undefined||grpSep===null||grpSep==='') ? ',' : String(grpSep).charAt(0);
  var s=String(text).trim();
  if(s==='') return 0;
  s=s.split(g).join('');
  if(d!=='.') s=s.split(d).join('.');
  var pct=0;
  while(/%\s*$/.test(s)){ pct++; s=s.replace(/%\s*$/,''); }
  if(!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(s)) return null;
  var n=parseFloat(s);
  return isNaN(n) ? null : n/Math.pow(100,pct);
}
function _jsDatestring(serial){var base=new Date(1899,11,30);var d=new Date(base.getTime()+serial*86400000);return d.getFullYear()+'/'+(d.getMonth()+1)+'/'+d.getDate();}

// --- 情報 ---
function _jsValue(x){var n=parseFloat(String(x).replace(/,/g,''));return isNaN(n)?null:n;}
function _jsN(v){if(typeof v==='number')return v;if(v===true)return 1;return 0;}
function _jsType(v){if(typeof v==='number')return 1;if(typeof v==='string')return 2;if(typeof v==='boolean')return 4;if(v&&v.type)return 16;if(Array.isArray(v))return 64;return 1;}
function _jsCell(infoType,val){if(infoType==='type'){if(val===''||val===null||val===undefined)return'b';if(typeof val==='number')return'n';return'l';}if(infoType==='contents')return val;return'';}

// --- 参照 ---
function _jsIndirect(sheet,refStr){var rc=_toRC(refStr.trim());if(rc&&_hf){var v=_hf.getCellValue({sheet:_hfSid(sheet),row:rc.r,col:rc.c});return v===null||v===undefined?'':v;}return'#REF!';}
function _jsOffset(sheet,ref,dr,dc){var rc=_toRC(ref.trim());if(!rc||!_hf)return'#REF!';var nr=rc.r+dr,nc=rc.c+dc;if(nr<0||nc<0)return'#REF!';var v=_hf.getCellValue({sheet:_hfSid(sheet),row:nr,col:nc});return v===null||v===undefined?'':v;}
// Excelの比較順: 数値 < 文字列 < 論理値。文字列は大小を区別しない。
function _xlCmp(a,b){
  var ra=(typeof a==='number')?0:(typeof a==='boolean'?2:1);
  var rb=(typeof b==='number')?0:(typeof b==='boolean'?2:1);
  if(ra!==rb) return ra<rb?-1:1;
  if(ra===0) return a<b?-1:(a>b?1:0);
  if(ra===2) return (a?1:0)-(b?1:0);
  var x=String(a).toLowerCase(), y=String(b).toLowerCase();
  return x<y?-1:(x>y?1:0);
}
// LOOKUP(ベクトル形式)。検索範囲は【昇順】が前提(Microsoftの仕様)。
//  見つからなければ「超えない最大」を返し、全部より小さければ #N/A。
//  ★昇順でない入力は Microsoft が「正しい値を返さないことがある」と明記している未定義動作。
//    その二分探索の実装差までは合わせない(合わせても意味がない)。
function _jsLookup(val,lookupVals,resultVals){
  var found=-1;
  for(var i=0;i<lookupVals.length;i++){
    var lv=lookupVals[i];
    if(lv===null||lv===undefined||lv==='') continue;
    if(_xlCmp(lv,val)<=0) found=i;
  }
  if(found<0) return '#N/A';
  var r=resultVals[found];
  return (r===undefined||r===null)?'':r;
}
function _jsXlookup(val,lookupVals,returnVals,notFound){for(var i=0;i<lookupVals.length;i++){if(String(lookupVals[i])===String(val))return returnVals[i];}return notFound!==undefined?notFound:'#N/A';}
// XMATCH(値, 配列, [一致モード], [検索モード])
//   一致モード: 0=完全一致(既定・MATCHの既定と違う) / -1=完全一致か次に小さい / 1=完全一致か次に大きい / 2=ワイルドカード
//   検索モード: 1=先頭から(既定) / -1=末尾から / 2,-2=二分探索(昇順前提。ここでは走査順としてだけ扱う)
function _jsXmatch(val,arr,matchMode,searchMode){
  matchMode = (matchMode===undefined||matchMode===null) ? 0 : Math.trunc(matchMode);
  searchMode = (searchMode===undefined||searchMode===null||searchMode===0) ? 1 : Math.trunc(searchMode);
  var order=[];
  for(var i=0;i<arr.length;i++) order.push(i);
  if(searchMode<0) order.reverse();
  if(matchMode===0 || matchMode===2){
    for(var j=0;j<order.length;j++){
      var k=order[j], v=arr[k];
      var hit = (matchMode===2 && typeof val==='string' && /[*?]/.test(val))
        ? _xlWildRe(val).test(String(v))
        : _xlCmp(v,val)===0;
      if(hit) return k+1;
    }
    return '#N/A';
  }
  // -1 = 超えない最大 / 1 = 下回らない最小
  var best=-1;
  for(var m=0;m<arr.length;m++){
    var w=arr[m];
    if(w===null||w===undefined||w==='') continue;
    var c=_xlCmp(w,val);
    if(matchMode===-1 && c<=0 && (best<0 || _xlCmp(w,arr[best])>0)) best=m;
    if(matchMode===1  && c>=0 && (best<0 || _xlCmp(w,arr[best])<0)) best=m;
  }
  return best<0 ? '#N/A' : best+1;
}
// Excelのワイルドカード(* ? ~エスケープ)を正規表現へ。大小は区別しない。
function _xlWildRe(pat){
  var out='';
  for(var i=0;i<pat.length;i++){
    var c=pat.charAt(i);
    if(c==='~' && i+1<pat.length){ out+=pat.charAt(i+1).replace(/[.*+?^${}()|[\]\\]/g,'\\$&'); i++; continue; }
    if(c==='*'){ out+='[\\s\\S]*'; continue; }
    if(c==='?'){ out+='[\\s\\S]'; continue; }
    out+=c.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  }
  return new RegExp('^'+out+'$','i');
}

// --- 動的配列 ---
function _jsSequence(rows,cols,start,step){start=start===undefined?1:start;step=step===undefined?1:step;var r=[],val=start;for(var i=0;i<rows*cols;i++){r.push(val);val+=step;}return r;}
function _jsRandarray(rows,cols,min,max,isInt){min=min||0;max=max||1;rows=rows||1;cols=cols||1;var r=[];for(var i=0;i<rows*cols;i++){var v=Math.random()*(max-min)+min;r.push(isInt?Math.floor(v):v);}return r;}
function _jsSort(vals,order){return vals.slice().sort(function(a,b){return order===false?b-a:a-b;});}
function _jsUnique(vals){return vals.filter(function(v,i,a){return a.indexOf(v)===i;});}
function _jsFilter(vals,include){return vals.filter(function(_,i){return!!include[i];});}
function _jsTake(arr,n){return n>0?arr.slice(0,n):arr.slice(arr.length+n);}
function _jsDrop(arr,n){return n>0?arr.slice(n):arr.slice(0,arr.length+n);}
function _jsExpand(arr,len,pad){var r=arr.slice();while(r.length<len)r.push(pad===undefined?0:pad);return r;}
function _jsVstack(){return[].concat.apply([],Array.from(arguments));}
function _jsHstack(){return[].concat.apply([],Array.from(arguments));}
function _jsTorow(arr){return[].concat(arr);}
function _jsTocol(arr){return[].concat(arr);}
function _jsChoosecols(matrix,cols){return cols.map(function(c){return matrix[c-1];});}
function _jsChooserows(matrix,rows){return rows.map(function(r){return matrix[r-1];});}

// --- 財務 ---
function _jsIrr(values,guess){guess=guess===undefined?0.1:guess;var rate=guess;for(var iter=0;iter<1000;iter++){var npv=0,dnpv=0;values.forEach(function(v,i){var f=Math.pow(1+rate,i);npv+=v/f;dnpv-=i*v/Math.pow(1+rate,i+1);});if(Math.abs(dnpv)<1e-15)break;var nr=rate-npv/dnpv;if(Math.abs(nr-rate)<1e-10){rate=nr;break;}rate=nr;}return rate;}
function _jsXirr(values,dates,guess){guess=guess===undefined?0.1:guess;var d0=dates[0],rate=guess;for(var iter=0;iter<1000;iter++){var npv=0,dnpv=0;values.forEach(function(v,i){var t=(dates[i]-d0)/365,f=Math.pow(1+rate,t);npv+=v/f;dnpv-=t*v/Math.pow(1+rate,t+1);});if(Math.abs(dnpv)<1e-15)break;var nr=rate-npv/dnpv;if(Math.abs(nr-rate)<1e-10){rate=nr;break;}rate=nr;}return rate;}
function _jsVdb(cost,salvage,life,start,end,factor){factor=factor||2;var total=0;for(var p=Math.floor(start);p<Math.ceil(end);p++){var s=Math.max(start,p),e2=Math.min(end,p+1);var book=cost;for(var i=0;i<p;i++)book-=Math.min(book*factor/life,book-salvage);total+=Math.min(book*factor/life,book-salvage)*(e2-s);}return total;}
function _jsDisc(settlement,maturity,pr,redemption){var t=(maturity-settlement)/365;return(redemption-pr)/(redemption*t);}
function _jsIntrate(settlement,maturity,investment,redemption){var t=(maturity-settlement)/365;return(redemption-investment)/(investment*t);}
function _jsDuration(settlement,maturity,coupon,yld,freq){var n=Math.round((maturity-settlement)/365*freq);var pv=0,wsum=0;for(var k=1;k<=n;k++){var cf=coupon*100/freq+(k===n?100:0);var pvk=cf/Math.pow(1+yld/freq,k);pv+=pvk;wsum+=pvk*k/freq;}return wsum/pv;}
function _jsAccrint(issue,firstInterest,settlement,rate,par,freq){var t=(settlement-issue)/365;return par*rate*t;}
function _jsAmordegrc(cost,date_purchased,first_period,salvage,period,rate){return cost*rate;}
function _jsAmorlinc(cost,date_purchased,first_period,salvage,period,rate){return(cost-salvage)*rate;}

// --- エンジニアリング ---
function _jsConvert(num,from,to){var u={'g':1,'kg':1000,'mg':0.001,'lbm':453.592,'ozm':28.3495,'ton':1e6,'m':1,'km':1000,'cm':0.01,'mm':0.001,'mi':1609.344,'yd':0.9144,'ft':0.3048,'in':0.0254,'sec':1,'min':60,'hr':3600,'day':86400,'yr':31557600,'Pa':1,'atm':101325,'bar':100000,'psi':6894.76,'J':1,'cal':4.184,'eV':1.602e-19,'l':0.001,'ml':0.000001,'gal':0.003785,'qt':0.000946,'pt':0.000473};if(from==='C'&&to==='F')return num*9/5+32;if(from==='F'&&to==='C')return(num-32)*5/9;if(from==='C'&&to==='K')return num+273.15;if(from==='K'&&to==='C')return num-273.15;if(from==='F'&&to==='K')return(num-32)*5/9+273.15;if(from==='K'&&to==='F')return(num-273.15)*9/5+32;if(!u[from]||!u[to])return'#N/A';return num*u[from]/u[to];}
function _jsGestep(num,step){return num>=(step||0)?1:0;}

// --- データベース ---
function _jsDbFunc(func,sheet,dbRange,field,criteriaRange){if(!_hf)return'#VALUE!';var m=dbRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);if(!m)return'#VALUE!';var s=_toRC(m[1]),e=_toRC(m[2]);var headers=[];for(var c=s.c;c<=e.c;c++)headers.push(_hf.getCellValue({sheet:_hfSid(sheet),row:s.r,col:c}));var fieldIdx=typeof field==='number'?field-1:headers.indexOf(field);if(fieldIdx<0)return'#VALUE!';var cm=criteriaRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);if(!cm)return'#VALUE!';var cs=_toRC(cm[1]);var critField=_hf.getCellValue({sheet:_hfSid(sheet),row:cs.r,col:cs.c});var critVal=_hf.getCellValue({sheet:_hfSid(sheet),row:cs.r+1,col:cs.c});var critFieldIdx=headers.indexOf(critField);var results=[];for(var r=s.r+1;r<=e.r;r++){if(String(_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:s.c+critFieldIdx}))===String(critVal)){results.push(_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:s.c+fieldIdx}));}}var nums=results.filter(function(v){return typeof v==='number';});if(func==='SUM')return nums.reduce(function(a,b){return a+b;},0);if(func==='AVG')return nums.length?nums.reduce(function(a,b){return a+b;})/nums.length:'#DIV/0!';if(func==='CNT')return nums.length;if(func==='CNTA')return results.filter(function(v){return v!==null&&v!==undefined;}).length;if(func==='MAX')return nums.length?Math.max.apply(null,nums):'#NUM!';if(func==='MIN')return nums.length?Math.min.apply(null,nums):'#NUM!';if(func==='PROD')return nums.reduce(function(a,b){return a*b;},1);if(func==='GET')return results.length===1?results[0]:'#NUM!';if(func==='STD'){var n=nums.length;if(n<2)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return Math.sqrt(nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/(n-1));}if(func==='STDP'){var n=nums.length;if(!n)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return Math.sqrt(nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/n);}if(func==='VAR'){var n=nums.length;if(n<2)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/(n-1);}if(func==='VARP'){var n=nums.length;if(!n)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/n;}return'#VALUE!';}

// --- Web ---
function _jsEncodeUrl(str){try{return encodeURIComponent(str);}catch(e){return'#VALUE!';}}

// --- AGGREGATE ---
function _jsAggregate(funcNum,opts,vals){var clean=vals.filter(function(v){return typeof v==='number'&&isFinite(v);});if(funcNum===1)return clean.reduce(function(a,b){return a+b;},0)/clean.length;if(funcNum===2)return clean.length;if(funcNum===3)return vals.filter(function(v){return v!==null&&v!==undefined;}).length;if(funcNum===4)return Math.max.apply(null,clean);if(funcNum===5)return Math.min.apply(null,clean);if(funcNum===6)return clean.reduce(function(a,b){return a*b;},1);if(funcNum===7){var n=clean.length;if(n<2)return'#DIV/0!';var m=clean.reduce(function(a,b){return a+b;})/n;return Math.sqrt(clean.reduce(function(a,v){return a+Math.pow(v-m,2);},0)/(n-1));}if(funcNum===8){var n=clean.length;if(!n)return'#DIV/0!';var m=clean.reduce(function(a,b){return a+b;})/n;return Math.sqrt(clean.reduce(function(a,v){return a+Math.pow(v-m,2);},0)/n);}if(funcNum===9)return clean.reduce(function(a,b){return a+b;},0);if(funcNum===10){var n=clean.length;if(n<2)return'#DIV/0!';var m=clean.reduce(function(a,b){return a+b;})/n;return clean.reduce(function(a,v){return a+Math.pow(v-m,2);},0)/(n-1);}if(funcNum===11){var n=clean.length;if(!n)return'#DIV/0!';var m=clean.reduce(function(a,b){return a+b;})/n;return clean.reduce(function(a,v){return a+Math.pow(v-m,2);},0)/n;}if(funcNum===12)return _jsMode(clean);if(funcNum===13){var s=clean.slice().sort(function(a,b){return a-b;});return s.length%2?s[Math.floor(s.length/2)]:(s[s.length/2-1]+s[s.length/2])/2;}return'#VALUE!';}

// ================================================================
// LET / LAMBDA / MAP / REDUCE / SCAN / MAKEARRAY / BYROW / BYCOL / ISOMITTED
// ================================================================

// 引数文字列をネスト考慮でパース
function _parseFuncArgs(str) {
  var args=[], depth=0, current='', inStr=false, strChar='';
  for(var i=0;i<str.length;i++){
    var ch=str[i];
    if(!inStr&&(ch==='"'||ch==="'")){inStr=true;strChar=ch;current+=ch;continue;}
    if(inStr&&ch===strChar){inStr=false;current+=ch;continue;}
    if(!inStr&&ch==='(')depth++;
    if(!inStr&&ch===')')depth--;
    if(!inStr&&ch===','&&depth===0){args.push(current.trim());current='';continue;}
    current+=ch;
  }
  if(current.trim())args.push(current.trim());
  return args;
}

// LET展開（再帰・セル参照名変数はスキップ）
function _jsLetExpand(formula) {
  if(!formula||formula[0]!=='=') return formula;
  var result = formula;
  for(var iter=0;iter<10;iter++){
    var f = result.slice(1).trim();
    if(!/^LET\s*\(/i.test(f)) break;
    var inside = f.match(/^LET\s*\((.+)\)$/is);
    if(!inside) break;
    var args = _parseFuncArgs(inside[1]);
    if(!args||args.length<3||args.length%2===0) break;
    var body = args[args.length-1];
    for(var i=args.length-3;i>=0;i-=2){
      var name=args[i].trim(), val=args[i+1].trim();
      if(/^[A-Z]+\d+$/i.test(name)) continue; // セル参照名の変数はスキップ
      if(/^\d/.test(name)) continue;
      body = body.replace(new RegExp('\\b'+name+'\\b','g'), '('+val+')');
    }
    result = '='+body;
  }
  return result;
}

// LAMBDA即時呼び出し展開
function _jsLambdaExpand(formula) {
  if(!formula||formula[0]!=='=') return formula;
  var f = formula.slice(1).trim();
  var mL = f.match(/^LAMBDA\s*\((.+)\)\s*\((.+)\)$/is);
  if(!mL) return formula;
  var defArgs = _parseFuncArgs(mL[1]);
  if(defArgs.length<2) return formula;
  var params=defArgs.slice(0,-1), body=defArgs[defArgs.length-1];
  var callArgs = _parseFuncArgs(mL[2]);
  if(callArgs.length!==params.length) return formula;
  params.forEach(function(p,i){
    var name=p.trim();
    if(/^[A-Z]+\d+$/i.test(name)) return;
    body = body.replace(new RegExp('\\b'+name+'\\b','g'),'('+callArgs[i].trim()+')');
  });
  return '='+body;
}

// REDUCE(initial, range, LAMBDA(acc, x, body)) → HF委譲
function _jsReduceCompute(sheet, initial, rangeStr, lambdaFormula) {
  if(!_hf) return null;
  var mL=lambdaFormula.match(/^LAMBDA\s*\(([^,)]+)\s*,\s*([^,)]+)\s*,\s*(.+)\)$/is);
  if(!mL) return '#VALUE!';
  var p1=mL[1].trim(),p2=mL[2].trim(),body=mL[3].trim();
  var m=rangeStr.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i); if(!m) return '#VALUE!';
  var s=_toRC(m[1]),e=_toRC(m[2]),acc=initial;
  for(var r=s.r;r<=e.r;r++){for(var c=s.c;c<=e.c;c++){
    var v=_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:c});
    var expr=body.replace(new RegExp('\\b'+p1+'\\b','g'),'('+acc+')').replace(new RegExp('\\b'+p2+'\\b','g'),'('+v+')');
    _hf.setCellContents({sheet:_hfSid(sheet),row:9998,col:0},'='+expr);
    var res=_hf.getCellValue({sheet:_hfSid(sheet),row:9998,col:0});
    if(typeof res==='number') acc=res;
    _hf.setCellContents({sheet:_hfSid(sheet),row:9998,col:0},null);
  }}
  return String(acc);
}

// SCAN(initial, range, LAMBDA(acc, x, body)) → 先頭値を返す
function _jsScanCompute(sheet, initial, rangeStr, lambdaFormula) {
  if(!_hf) return null;
  var mL=lambdaFormula.match(/^LAMBDA\s*\(([^,)]+)\s*,\s*([^,)]+)\s*,\s*(.+)\)$/is);
  if(!mL) return '#VALUE!';
  var p1=mL[1].trim(),p2=mL[2].trim(),body=mL[3].trim();
  var m=rangeStr.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i); if(!m) return '#VALUE!';
  var s=_toRC(m[1]),e=_toRC(m[2]),acc=initial,first=null;
  for(var r=s.r;r<=e.r;r++){for(var c=s.c;c<=e.c;c++){
    var v=_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:c});
    var expr=body.replace(new RegExp('\\b'+p1+'\\b','g'),'('+acc+')').replace(new RegExp('\\b'+p2+'\\b','g'),'('+v+')');
    _hf.setCellContents({sheet:_hfSid(sheet),row:9997,col:0},'='+expr);
    var res=_hf.getCellValue({sheet:_hfSid(sheet),row:9997,col:0});
    if(typeof res==='number'){acc=res; if(first===null)first=res;}
    _hf.setCellContents({sheet:_hfSid(sheet),row:9997,col:0},null);
  }}
  return first!==null ? String(first) : String(initial);
}

// MAP(range, LAMBDA(x, body)) → 全件カンマ区切り
function _jsMapCompute(sheet, rangeStr, lambdaFormula) {
  if(!_hf) return null;
  var mL=lambdaFormula.match(/^LAMBDA\s*\(([^,)]+)\s*,\s*(.+)\)$/is);
  if(!mL) return '#VALUE!';
  var param=mL[1].trim(),body=mL[2].trim();
  var m=rangeStr.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i); if(!m) return '#VALUE!';
  var s=_toRC(m[1]),e=_toRC(m[2]),results=[];
  for(var r=s.r;r<=e.r;r++){for(var c=s.c;c<=e.c;c++){
    var v=_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:c});
    var expr=body.replace(new RegExp('\\b'+param+'\\b','g'),'('+v+')');
    _hf.setCellContents({sheet:_hfSid(sheet),row:9999,col:c},'='+expr);
    var res=_hf.getCellValue({sheet:_hfSid(sheet),row:9999,col:c});
    results.push(res===null?'':String(res));
    _hf.setCellContents({sheet:_hfSid(sheet),row:9999,col:c},null);
  }}
  return results[0]; // 先頭値（スピル制限）
}

// MAKEARRAY(rows, cols, LAMBDA(r, c, body)) → 先頭値を返す
function _jsMakearrayCompute(sheet, rows, cols, lambdaFormula) {
  if(!_hf) return null;
  var mL=lambdaFormula.match(/^LAMBDA\s*\(([^,)]+)\s*,\s*([^,)]+)\s*,\s*(.+)\)$/is);
  if(!mL) return '#VALUE!';
  var p1=mL[1].trim(),p2=mL[2].trim(),body=mL[3].trim();
  var expr=body.replace(new RegExp('\\b'+p1+'\\b','g'),'(1)').replace(new RegExp('\\b'+p2+'\\b','g'),'(1)');
  _hf.setCellContents({sheet:_hfSid(sheet),row:9996,col:0},'='+expr);
  var res=_hf.getCellValue({sheet:_hfSid(sheet),row:9996,col:0});
  _hf.setCellContents({sheet:_hfSid(sheet),row:9996,col:0},null);
  return res===null?'0':String(res);
}

// ISOMITTED(arg) - 省略チェック
function _jsIsomitted(v) { return v===undefined||v===null; }

// ================================================================
// convertFormula: 文字列変換
// ================================================================
// VALUE("1,234") を Excel と同じ 1234 にする。
//  ★独自層(_jsComputeFormula)は「式の一番外側の関数」しか横取りしないため、
//    =IFERROR(VALUE(B7),"NA") のように入れ子だと届かない。
//    そこで文字列書換の段(convertFormula)で桁区切りを外す形に書き換える。ここなら入れ子でも効く。
//  ・文字列リテラルの中は書き換えない
//  ・DATEVALUE / NUMBERVALUE のように VALUE で終わる別の関数名は触らない(直前が識別子の文字なら飛ばす)
//  関数呼び出しを丸ごと別の式に置き換える汎用。文字列リテラルの中は触らない。
//  DATEVALUE / NUMBERVALUE や POINT( のように名前の一部として現れる物は触らない(直前が識別子の文字なら飛ばす)。
function _rewriteCalls(f, name, build) {
  var out='', i=0, ins=false;
  var re = new RegExp('^' + name + '\\s*\\(', 'i');
  while(i<f.length){
    var ch=f.charAt(i);
    if(ch==='"'){ ins=!ins; out+=ch; i++; continue; }
    if(ins){ out+=ch; i++; continue; }
    var m=re.exec(f.slice(i));
    var prev=i>0?f.charAt(i-1):'';
    if(m && !/[A-Za-z0-9_.]/.test(prev)){
      var start=i+m[0].length, depth=1, j=start, sin=false;
      for(; j<f.length; j++){
        var c=f.charAt(j);
        if(c==='"'){ sin=!sin; continue; }
        if(sin) continue;
        if(c==='(') depth++;
        else if(c===')'){ depth--; if(depth===0) break; }
      }
      if(depth!==0){ out+=ch; i++; continue; }          // 括弧が閉じていない=触らない
      var inner=_rewriteCalls(f.slice(start,j), name, build);   // 入れ子も先に置き換える
      var built=build(_parseFuncArgs(inner));
      if(built===null){ out+=f.slice(i, j+1); i=j+1; continue; }
      out += built;
      i = j+1;
      continue;
    }
    out+=ch; i++;
  }
  return out;
}
// floor(x) を HyperFormula の式で書く。HFのINTは0方向へ切り捨てるので使えない。
function _floorExpr(x){ return '(IF((' + x + ')<0,-ROUNDUP(-(' + x + '),0),ROUNDDOWN((' + x + '),0)))'; }

function _rewriteValueCalls(f) {
  // ★HyperFormula 2.6.1 には VALUE 関数自体が無い(#NAME?)。VALUE のまま渡してもだめなので、
  //   「桁区切りを外して数値に変える」形に置き換える。数値にならない物は #VALUE! になり、
  //   Excel と同じく IFERROR で拾える。空文字は Excel と同じく #VALUE! にする("x"*1 でエラーを作る)。
  return _rewriteCalls(f, 'VALUE', function(a){
    if(a.length!==1) return null;
    var x=a[0];
    return '((IF((' + x + ')="","x",SUBSTITUTE((' + x + '),",","")))*1)';
  });
}
// ★独自層(_jsComputeFormula)は「式の一番外側の関数」しか横取りしないので、
//   =ROUND(MOD(-3,2),0) のように入れ子だと届かず、HFの違う答えがそのまま出ていた。
//   INT と MOD は HF の式で正しく書けるので、ここで書き換えて入れ子でも合うようにする。
//   (TEXT は書式処理を式で書けないため、ここでは直せない=台帳 R12 で管理)
function _rewriteIntMod(f) {
  f = _rewriteCalls(f, 'INT', function(a){ return a.length===1 ? _floorExpr(a[0]) : null; });
  f = _rewriteCalls(f, 'MOD', function(a){
    if(a.length!==2) return null;
    var x=a[0], y=a[1], q='((' + x + ')/(' + y + '))';
    return '((' + x + ')-(' + y + ')*' + _floorExpr(q) + ')';
  });
  return f;
}

function convertFormula(f) {
  if(!f||f[0]!=='=') return f;
  if(/VALUE\s*\(/i.test(f))     f = _rewriteValueCalls(f);
  if(/\b(INT|MOD)\s*\(/i.test(f)) f = _rewriteIntMod(f);
  // LET展開
  if(/^=LET\s*\(/i.test(f)) f = _jsLetExpand(f);
  // LAMBDA即時呼び出し展開
  if(/^=LAMBDA\s*\(/i.test(f)) f = _jsLambdaExpand(f);
  //  ★JIS( → DBCS( … 半角→全角の関数の本名は DBCS。JIS は日本語UIの表示名で、
  //    xlsx の中身も US-English構文も DBCS。日本語Excelの癖で JIS と打つ人が居るので、
  //    ここ(入口)で本名に直す＝エンジンにも書き出しにも DBCS しか流れない
  //    (JIS のまま書き出すと Excel がその式を #NAME? にする。実Excelで確認済み)。
  f = f.replace(/\bJIS\s*\(/gi, 'DBCS(');
  f = f.replace(/\bFALSE\b(?!\s*\()/g, 'FALSE()');
  f = f.replace(/\bTRUE\b(?!\s*\()/g, 'TRUE()');
  f = f.replace(/\bNORMSDIST\s*\(([^,)]+)\)/gi, 'NORMSDIST($1,TRUE())');
  f = f.replace(/\bBETADIST\s*\(([^,()]+,[^,()]+,[^,()]+)\)/gi, 'BETA.DIST($1,TRUE())');
  f = f.replace(/\bHYPGEOMDIST\s*\(([^()]+)\)/gi, function(m,args){return args.split(',').length===4?'HYPGEOMDIST('+args+',FALSE())':m;});
  f = f.replace(/\bNEGBINOMDIST\s*\(([^()]+)\)/gi, function(m,args){return args.split(',').length===3?'NEGBINOMDIST('+args+',FALSE())':m;});
  f = f.replace(/\bISREF\s*\(([^)]+)\)/gi, function(m,arg){return /^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(arg.trim())?'TRUE()':'FALSE()';});
  return f;
}

// ================================================================
// _jsComputeFormula: JS側で完全計算
// ================================================================
function _jsComputeFormula(sheet, v) {
  if(!v||v[0]!=='='||!_hf) return null;

  // 🚀 早期return: 関数名でJS処理対象外をHFに早送り（55regex→O(1)）
  var _f0 = v.slice(1).trim();
  var _fn = _f0.match(/^([A-Z][A-Z0-9.]*)\s*\(/i);
  if(!_fn) return null; // 関数形式でない（=A1+B1等）→ HFへ
  var _fnBase = _fn[1].toUpperCase().split('.')[0];
  //  ★第3波P1(2026-08-01)で 11関数をここから外して HFプラグインへ移した:
  //    CONCAT / LOOKUP / XMATCH / INDIRECT / DATEVALUE / NUMBERVALUE / FIXED / ASC / DBCS(JIS) /
  //    TEXTBEFORE / TEXTAFTER
  //    理由=ここに居ると「式の一番外側」でしか効かない。=ROUND(LOOKUP(...),0) のような入れ子で
  //    HFの答え(多くは #NAME?)に戻り、表示だけ正しくて参照先が間違う。
  //  ★1つの関数は1箇所でだけ定義する（両方に居たら compare.mjs が赤にする）。
  var _jsSet = {RANK:1,PERCENTILE:1,QUARTILE:1,
    MODE:1,TRIMMEAN:1,PERCENTRANK:1,KURT:1,INTERCEPT:1,FORECAST:1,IRR:1,XIRR:1,
    DATESTRING:1,OFFSET:1,
    DOLLAR:1,YEN:1,N:1,TYPE:1,ENCODEURL:1,
    CONVERT:1,GESTEP:1,VALUETOTEXT:1,
    DSUM:1,DAVERAGE:1,DCOUNT:1,DCOUNTA:1,DMAX:1,DMIN:1,DPRODUCT:1,
    DGET:1,DSTDEV:1,DSTDEVP:1,DVAR:1,DVARP:1,
    AGGREGATE:1,LINEST:1,PERMUT:1,PERMUTATIONA:1,BINOM:1,FREQUENCY:1,
    MDETERM:1,REDUCE:1,SCAN:1,MAP:1,MAKEARRAY:1,ISOMITTED:1,
    LENB:1,LEFTB:1,RIGHTB:1,MIDB:1};
  if(!_jsSet[_fnBase]) return null; // JS非対象 → HFへ

  var f = v.slice(1).trim().toUpperCase();
  var fOrig = v.slice(1).trim();

  // RANK / RANK.EQ / RANK.AVG
  var mRank=fOrig.match(/^RANK(?:\.EQ|\.AVG)?\s*\(([A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([01])\s*\)$/i);
  if(mRank){var val=_getSingleVal(sheet,mRank[1]);var vals=_getRangeVals(sheet,mRank[2]);if(val!==null&&vals.length)return String(_jsRank(val,vals,parseInt(mRank[3])));}

  // PERCENTILE / PERCENTILE.INC / PERCENTILE.EXC
  var mPerc=fOrig.match(/^PERCENTILE(?:\.INC|\.EXC)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([0-9.]+)\s*\)$/i);
  if(mPerc){var vals=_getRangeVals(sheet,mPerc[1]);if(vals.length)return String(_jsPercentile(vals,parseFloat(mPerc[2])));}

  // QUARTILE / QUARTILE.INC / QUARTILE.EXC
  var mQuart=fOrig.match(/^QUARTILE(?:\.INC|\.EXC)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([0-4])\s*\)$/i);
  if(mQuart){var vals=_getRangeVals(sheet,mQuart[1]);if(vals.length)return String(_jsQuartile(vals,parseInt(mQuart[2])));}

  // MODE / MODE.SNGL / MODE.MULT
  var mMode=fOrig.match(/^MODE(?:\.SNGL|\.MULT)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mMode){var vals=_getRangeVals(sheet,mMode[1]);return String(_jsMode(vals));}

  // TRIMMEAN
  var mTrim=fOrig.match(/^TRIMMEAN\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([0-9.]+)\s*\)$/i);
  if(mTrim){var vals=_getRangeVals(sheet,mTrim[1]);return String(_jsTrimmean(vals,parseFloat(mTrim[2])));}

  // PERCENTRANK / PERCENTRANK.INC / PERCENTRANK.EXC
  var mPr=fOrig.match(/^PERCENTRANK(?:\.INC|\.EXC)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+|[0-9.]+)\s*(?:,\s*(\d+))?\s*\)$/i);
  if(mPr){var vals=_getRangeVals(sheet,mPr[1]);var x=_getSingleVal(sheet,mPr[2])||parseFloat(mPr[2]);return String(_jsPercentrank(vals,x,parseInt(mPr[3])||3));}

  // KURT
  var mKurt=fOrig.match(/^KURT\s*\(([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mKurt){var vals=_getRangeVals(sheet,mKurt[1]);var r=_jsKurt(vals);return typeof r==='number'?String(Math.round(r*10000)/10000):r;}

  // INTERCEPT
  var mInter=fOrig.match(/^INTERCEPT\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mInter){var ys=_getRangeVals(sheet,mInter[1]),xs=_getRangeVals(sheet,mInter[2]);return String(_jsIntercept(ys,xs));}

  // FORECAST / FORECAST.LINEAR
  var mFc=fOrig.match(/^FORECAST(?:\.LINEAR)?\s*\(([A-Z]+\d+|[0-9.]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mFc){var x=_getSingleVal(sheet,mFc[1])||parseFloat(mFc[1]);var ys=_getRangeVals(sheet,mFc[2]),xs=_getRangeVals(sheet,mFc[3]);return String(_jsForecast(x,ys,xs));}

  // IRR
  var mIrr=fOrig.match(/^IRR\s*\(([A-Z]+\d+:[A-Z]+\d+)(?:\s*,\s*[0-9.]+)?\s*\)$/i);
  if(mIrr){var vals=_getRangeAll(sheet,mIrr[1]).filter(function(v){return typeof v==='number';});return String(Math.round(_jsIrr(vals)*10000)/10000);}

  // DATESTRING
  var mDs=fOrig.match(/^DATESTRING\s*\(([A-Z]+\d+|[0-9]+)\)$/i);
  if(mDs){var sv=_getSingleVal(sheet,mDs[1])||parseInt(mDs[1]);return _jsDatestring(sv);}

  // OFFSET
  var mOff=fOrig.match(/^OFFSET\s*\(([A-Z]+\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)$/i);
  if(mOff){var v=_jsOffset(sheet,mOff[1],parseInt(mOff[2]),parseInt(mOff[3]));return String(v);}

  // XLOOKUP
  var mXl=fOrig.match(/^XLOOKUP\s*\(([^,()]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)(?:\s*,\s*"([^"]*)")?\s*\)$/i);
  if(mXl){var lv=_getSingleVal(sheet,mXl[1]);if(lv===null)lv=mXl[1].replace(/^["']|["']$/g,'');var la=_getRangeAll(sheet,mXl[2]),ra=_getRangeAll(sheet,mXl[3]);return String(_jsXlookup(lv,la,ra,mXl[4]));}

  // DOLLAR
  var mDollar=fOrig.match(/^DOLLAR\s*\(([A-Z]+\d+|[0-9.\-]+)(?:\s*,\s*(-?\d+))?\s*\)$/i);
  if(mDollar){var n=_getSingleVal(sheet,mDollar[1])||parseFloat(mDollar[1]);return _jsDollar(n,mDollar[2]?parseInt(mDollar[2]):2);}

  // YEN
  var mYen=fOrig.match(/^YEN\s*\(([A-Z]+\d+|[0-9.\-]+)(?:\s*,\s*(\d+))?\s*\)$/i);
  if(mYen){var n=_getSingleVal(sheet,mYen[1])||parseFloat(mYen[1]);return _jsYen(n,mYen[2]?parseInt(mYen[2]):0);}

  // N
  var mN=fOrig.match(/^N\s*\(([^)]+)\)$/i);
  if(mN){var sv=_getSingleVal(sheet,mN[1]);return String(_jsN(sv!==null?sv:mN[1]));}

  // TYPE
  var mType=fOrig.match(/^TYPE\s*\(([^)]+)\)$/i);
  if(mType){var sv=_getSingleVal(sheet,mType[1]);return String(_jsType(sv!==null?sv:mType[1]));}

  // ENCODEURL
  var mEnc=fOrig.match(/^ENCODEURL\s*\(([^)]+)\)$/i);
  if(mEnc){var sv=_getSingleVal(sheet,mEnc[1])||mEnc[1].replace(/^["']|["']$/g,'');return _jsEncodeUrl(String(sv));}

  // CONVERT
  var mCv=fOrig.match(/^CONVERT\s*\(([A-Z]+\d+|[0-9.\-]+)\s*,\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)$/i);
  if(mCv){var n=_getSingleVal(sheet,mCv[1])||parseFloat(mCv[1]);var r=_jsConvert(n,mCv[2],mCv[3]);return typeof r==='number'?String(Math.round(r*10000)/10000):r;}

  // GESTEP
  var mGe=fOrig.match(/^GESTEP\s*\(([A-Z]+\d+|[0-9.\-]+)(?:\s*,\s*([A-Z]+\d+|[0-9.\-]+))?\s*\)$/i);
  if(mGe){var n=_getSingleVal(sheet,mGe[1])||parseFloat(mGe[1]);var s=mGe[2]?(_getSingleVal(sheet,mGe[2])||parseFloat(mGe[2])):0;return String(_jsGestep(n,s));}

  // VALUETOTEXT
  var mVtt=fOrig.match(/^VALUETOTEXT\s*\(([^)]+)\)$/i);
  if(mVtt){var sv=_getSingleVal(sheet,mVtt[1]);return _jsValuetotext(sv);}

  // REDUCE(initial, range, LAMBDA(...))
  var mReduce=fOrig.match(/^REDUCE\s*\(([^,]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*(LAMBDA\s*\(.+\))\s*\)$/is);
  if(mReduce){var reduceInit=parseFloat(mReduce[1].trim())||0;var reduceR=_jsReduceCompute(sheet,reduceInit,mReduce[2],mReduce[3]);if(reduceR!==null)return reduceR;}

  // SCAN(initial, range, LAMBDA(...))
  var mScan=fOrig.match(/^SCAN\s*\(([^,]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*(LAMBDA\s*\(.+\))\s*\)$/is);
  if(mScan){var scanInit=parseFloat(mScan[1].trim())||0;var scanR=_jsScanCompute(sheet,scanInit,mScan[2],mScan[3]);if(scanR!==null)return scanR;}

  // MAP(range, LAMBDA(...))
  var mMap=fOrig.match(/^MAP\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*(LAMBDA\s*\(.+\))\s*\)$/is);
  if(mMap){var mapR=_jsMapCompute(sheet,mMap[1],mMap[2]);if(mapR!==null)return mapR;}

  // MAKEARRAY(rows, cols, LAMBDA(...))
  var mMkarr=fOrig.match(/^MAKEARRAY\s*\(([0-9]+)\s*,\s*([0-9]+)\s*,\s*(LAMBDA\s*\(.+\))\s*\)$/is);
  if(mMkarr){var mkR=_jsMakearrayCompute(sheet,parseInt(mMkarr[1]),parseInt(mMkarr[2]),mMkarr[3]);if(mkR!==null)return mkR;}

  // ISOMITTED(ref)
  var mIso=fOrig.match(/^ISOMITTED\s*\(([^)]*)\)$/i);
  if(mIso){var isoSv=_getSingleVal(sheet,mIso[1].trim());return String(_jsIsomitted(isoSv));}

  // DSUM / DAVERAGE / DCOUNT / DCOUNTA / DMAX / DMIN / DPRODUCT / DGET / DSTDEV / DSTDEVP / DVAR / DVARP
  var mDb=fOrig.match(/^(DSUM|DAVERAGE|DCOUNT|DCOUNTA|DMAX|DMIN|DPRODUCT|DGET|DSTDEV|DSTDEVP|DVAR|DVARP)\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([^,]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*\)$/i);
  if(mDb){var funcMap={'DSUM':'SUM','DAVERAGE':'AVG','DCOUNT':'CNT','DCOUNTA':'CNTA','DMAX':'MAX','DMIN':'MIN','DPRODUCT':'PROD','DGET':'GET','DSTDEV':'STD','DSTDEVP':'STDP','DVAR':'VAR','DVARP':'VARP'};var fn=funcMap[mDb[1].toUpperCase()];var fieldArg=mDb[3].trim().replace(/^["']|["']$/g,'');var fieldVal=parseInt(fieldArg)||fieldArg;return String(_jsDbFunc(fn,sheet,mDb[2],fieldVal,mDb[4]));}

  // AGGREGATE
  var mAgg=fOrig.match(/^AGGREGATE\s*\((\d+)\s*,\s*(\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*\)$/i);
  if(mAgg){var vals=_getRangeVals(sheet,mAgg[3]);var r=_jsAggregate(parseInt(mAgg[1]),parseInt(mAgg[2]),vals);return String(typeof r==='number'?Math.round(r*10000)/10000:r);}

  // LINEST
  var mLi=fOrig.match(/^LINEST\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*\)$/i);
  if(mLi){var ys=_getRangeVals(sheet,mLi[1]),xs=_getRangeVals(sheet,mLi[2]);var r=_jsLinest(ys,xs);return String(Math.round(r[0]*10000)/10000);}

  // PERMUT
  var mPerm=fOrig.match(/^PERMUT\s*\(([0-9]+)\s*,\s*([0-9]+)\s*\)$/i);
  if(mPerm)return String(_jsPermut(parseInt(mPerm[1]),parseInt(mPerm[2])));

  // PERMUTATIONA
  var mPa=fOrig.match(/^PERMUTATIONA\s*\(([0-9]+)\s*,\s*([0-9]+)\s*\)$/i);
  if(mPa)return String(_jsPermutationa(parseInt(mPa[1]),parseInt(mPa[2])));

  // BINOM.DIST.RANGE
  var mBdr=fOrig.match(/^BINOM\.DIST\.RANGE\s*\(([0-9]+)\s*,\s*([0-9.]+)\s*,\s*([0-9]+)(?:\s*,\s*([0-9]+))?\s*\)$/i);
  if(mBdr)return String(Math.round(_jsBinomDistRange(parseInt(mBdr[1]),parseFloat(mBdr[2]),parseInt(mBdr[3]),mBdr[4]?parseInt(mBdr[4]):undefined)*10000)/10000);

  // FREQUENCY - 最初の値だけ返す（スピル）
  var mFreq=fOrig.match(/^FREQUENCY\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mFreq){var data=_getRangeVals(sheet,mFreq[1]),bins=_getRangeVals(sheet,mFreq[2]);var r=_jsFrequency(data,bins);return String(r[0]);}

  // MDETERM
  var mMd=fOrig.match(/^MDETERM\s*\(([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mMd){var m2=mMd[1].match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);if(m2){var s=_toRC(m2[1]),e=_toRC(m2[2]),n=e.r-s.r+1,mat=[];for(var r=0;r<n;r++){mat.push([]);for(var c=0;c<n;c++)mat[r].push(_hf.getCellValue({sheet:_hfSid(sheet),row:s.r+r,col:s.c+c})||0);}return String(Math.round(_jsMdeterm(mat)*10000)/10000);}}

  // IRR / XIRR (既存)
  var mXirr=fOrig.match(/^XIRR\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mXirr){var vals=_getRangeAll(sheet,mXirr[1]).filter(function(v){return typeof v==='number';});var dates=_getRangeAll(sheet,mXirr[2]).filter(function(v){return typeof v==='number';});return String(Math.round(_jsXirr(vals,dates)*10000)/10000);}

  // ENCODEURL (文字列直接)
  var mEncStr=fOrig.match(/^ENCODEURL\s*\("([^"]*)"\)$/i);
  if(mEncStr)return _jsEncodeUrl(mEncStr[1]);

  // LENB/LEFTB/RIGHTB/MIDB/FINDB
  var mLenb=fOrig.match(/^LENB\s*\(([A-Z]+\d+|"[^"]+")\)$/i);
  if(mLenb){var sv=_getSingleVal(sheet,mLenb[1])||mLenb[1].replace(/^"|"$/g,'');return String(_jsLenb(String(sv)));}
  var mLeftb=fOrig.match(/^LEFTB\s*\(([A-Z]+\d+|"[^"]+")\s*,\s*([0-9]+)\)$/i);
  if(mLeftb){var sv=_getSingleVal(sheet,mLeftb[1])||mLeftb[1].replace(/^"|"$/g,'');return _jsLeftb(String(sv),parseInt(mLeftb[2]));}
  var mRightb=fOrig.match(/^RIGHTB\s*\(([A-Z]+\d+|"[^"]+")\s*,\s*([0-9]+)\)$/i);
  if(mRightb){var sv=_getSingleVal(sheet,mRightb[1])||mRightb[1].replace(/^"|"$/g,'');return _jsRightb(String(sv),parseInt(mRightb[2]));}
  var mMidb=fOrig.match(/^MIDB\s*\(([A-Z]+\d+|"[^"]+")\s*,\s*([0-9]+)\s*,\s*([0-9]+)\)$/i);
  if(mMidb){var sv=_getSingleVal(sheet,mMidb[1])||mMidb[1].replace(/^"|"$/g,'');return _jsMidb(String(sv),parseInt(mMidb[2]),parseInt(mMidb[3]));}

  // DATESTRING
  var mDstr=fOrig.match(/^DATESTRING\s*\(([A-Z]+\d+|[0-9]+)\)$/i);
  if(mDstr){var sv=_getSingleVal(sheet,mDstr[1])||parseInt(mDstr[1]);return _jsDatestring(sv);}

  return null;
}

// ================================================================
// setCellFormula / undoCellFormula / recalcSheet
// ================================================================
// HFに渡す値の型変換（文字列の数値→number、空→null）

// ================================================================
// 【Node.js 両対応】module.exports
// ブラウザでは module が undefined なので無視される
// ================================================================
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    initExallyFormula: initExallyFormula,
    // 補助・内部
    _hfSid: _hfSid, _toRC: _toRC, _getRangeVals: _getRangeVals,
    _getRangeAll: _getRangeAll, _getSingleVal: _getSingleVal,
    _hfGetDisplay: _hfGetDisplay, _applyTextFormat: _applyTextFormat,
    addSheetToEngine: addSheetToEngine,
    // 統計
    _jsRank: _jsRank, _jsPercentile: _jsPercentile, _jsQuartile: _jsQuartile,
    _jsMode: _jsMode, _jsTrimmean: _jsTrimmean, _jsPercentrank: _jsPercentrank,
    _jsFrequency: _jsFrequency, _jsKurt: _jsKurt, _jsPermut: _jsPermut,
    _jsPermutationa: _jsPermutationa, _jsProb: _jsProb, _jsBinomDistRange: _jsBinomDistRange,
    // 回帰
    _jsSlope: _jsSlope, _jsIntercept: _jsIntercept, _jsForecast: _jsForecast,
    _jsLinest: _jsLinest, _jsLogest: _jsLogest, _jsTrend: _jsTrend, _jsGrowth: _jsGrowth,
    // 行列
    _jsMdeterm: _jsMdeterm, _jsMinverse: _jsMinverse,
    // 文字列
    _jsConcat: _jsConcat, _jsTextjoin: _jsTextjoin, _jsFixed: _jsFixed,
    _jsDollar: _jsDollar, _jsYen: _jsYen, _jsLenb: _jsLenb, _jsLeftb: _jsLeftb,
    _jsRightb: _jsRightb, _jsMidb: _jsMidb, _jsFindb: _jsFindb, _jsReplaceb: _jsReplaceb,
    _jsAsc: _jsAsc, _jsDbcs: _jsDbcs, _jsJis: _jsJis,
    _jsTextbefore: _jsTextbefore, _jsTextafter: _jsTextafter, _jsTextsplit: _jsTextsplit,
    _jsValuetotext: _jsValuetotext, _jsArraytotext: _jsArraytotext,
    // 日付・変換
    _jsDateValue: _jsDateValue, _jsDatestring: _jsDatestring,
    _jsValue: _jsValue, _jsNumbervalue: _jsNumbervalue, _jsN: _jsN, _jsType: _jsType, _jsCell: _jsCell,
    _xlRound: _xlRound, _xlCmp: _xlCmp, _xlWildRe: _xlWildRe,
    // 参照系
    _jsIndirect: _jsIndirect, _jsOffset: _jsOffset,
    _jsLookup: _jsLookup, _jsXlookup: _jsXlookup, _jsXmatch: _jsXmatch,
    // 配列
    _jsSequence: _jsSequence, _jsRandarray: _jsRandarray, _jsSort: _jsSort,
    _jsUnique: _jsUnique, _jsFilter: _jsFilter, _jsTake: _jsTake, _jsDrop: _jsDrop,
    _jsExpand: _jsExpand, _jsVstack: _jsVstack, _jsHstack: _jsHstack,
    _jsTorow: _jsTorow, _jsTocol: _jsTocol,
    _jsChoosecols: _jsChoosecols, _jsChooserows: _jsChooserows,
    // 財務
    _jsIrr: _jsIrr, _jsXirr: _jsXirr, _jsVdb: _jsVdb, _jsDisc: _jsDisc,
    _jsIntrate: _jsIntrate, _jsDuration: _jsDuration, _jsAccrint: _jsAccrint,
    _jsAmordegrc: _jsAmordegrc, _jsAmorlinc: _jsAmorlinc,
    // その他
    _jsConvert: _jsConvert, _jsGestep: _jsGestep, _jsDbFunc: _jsDbFunc,
    _jsEncodeUrl: _jsEncodeUrl, _jsAggregate: _jsAggregate,
    _jsIsomitted: _jsIsomitted,
    // LET/LAMBDA系
    _parseFuncArgs: _parseFuncArgs,
    _jsLetExpand: _jsLetExpand, _jsLambdaExpand: _jsLambdaExpand,
    _jsReduceCompute: _jsReduceCompute, _jsScanCompute: _jsScanCompute,
    _jsMapCompute: _jsMapCompute, _jsMakearrayCompute: _jsMakearrayCompute,
    // ルーター
    convertFormula: convertFormula,
    _jsComputeFormula: _jsComputeFormula
  };
}

// ================================================================
// 【HyperFormula 関数プラグイン】★ここが「その関数の唯一の定義場所」
// ================================================================
//  なぜプラグインなのか:
//    独自層(_jsComputeFormula)は「式の一番外側の関数」しか横取りしない。
//    =INDEX(SORT(...),1) や =IF(1=1,XLOOKUP(...),...) のように入れ子で使われると
//    救済が効かず、HyperFormula の違う答えがそのまま出る(実測で275本中61本)。
//    表示だけ正しくて、そのセルを参照した先が間違う=一番タチが悪い壊れ方。
//    プラグインとして登録すればエンジン自身が正しく計算するので、入れ子でも参照先でも正しくなる。
//
//  ★決まり: ここに登録した関数は _jsSet に入れない(1つの関数は1箇所でだけ定義する)。
//           テストが両方に居たら赤にする。
//  ★振り分けの基準:
//     ・入れ子で使われうる純粋な関数              → ここ(HFプラグイン)
//     ・セルを直接読む必要があるグリッド固有のオペ → _jsSet
var _PLUGIN_FUNCS = ['SORT','UNIQUE','TEXT','TEXTJOIN','INT','MOD','VALUE','XLOOKUP','FILTER','MATCH','SUMPRODUCT',
  // ★第3波P1(2026-08-01)
  'CONCAT','LOOKUP','XMATCH','INDIRECT','DATEVALUE','NUMBERVALUE','FIXED','ASC','DBCS','TEXTBEFORE','TEXTAFTER'];
var _pluginRegistered = false;

// 渡された物がクラス(HyperFormula)でも名前空間({HyperFormula, FunctionPlugin, ...})でも動くようにする。
//  ブラウザとNodeで形が違うため。
function _resolveHFParts(ns) {
  if(!ns) return null;
  var cls = (typeof ns.registerFunctionPlugin === 'function') ? ns
          : (ns.HyperFormula && typeof ns.HyperFormula.registerFunctionPlugin === 'function') ? ns.HyperFormula
          : null;
  if(!cls) return null;
  var g = (typeof window !== 'undefined' && window) ? window : {};
  function pick(k){ return ns[k] || cls[k] || g[k]; }
  return {
    cls: cls,
    FunctionPlugin: pick('FunctionPlugin'),
    T: pick('FunctionArgumentType'),
    SRV: pick('SimpleRangeValue'),
    CellError: pick('CellError'),
    ErrorType: pick('ErrorType')
  };
}

function registerExallyFunctions(HFns) {
  if(_pluginRegistered) return true;
  var P = _resolveHFParts(HFns);
  if(!P) return false;
  var FunctionPlugin = P.FunctionPlugin;
  var T = P.T;
  var SRV = P.SRV;
  var CellError = P.CellError, ErrorType = P.ErrorType;
  if(!FunctionPlugin || !T || !SRV || !CellError || !ErrorType) return false;

  function err(t){ return new CellError(t); }
  //  HyperFormula は空セルを Symbol(Empty value) で渡してくる。文字列化すると
  //  「Symbol(Empty value)」という文字がそのまま出るので、空文字に直す(実測で踏んだ)。
  function unwrap(v){ return (typeof v === 'symbol') ? '' : v; }
  function flat(v){
    if(v && typeof v==='object' && v.data) return [].concat.apply([], v.data).map(unwrap);
    if(Array.isArray(v)) return [].concat.apply([], v).map(unwrap);
    return [unwrap(v)];
  }
  function isErr(v){ return v instanceof CellError; }
  function firstErr(arr){ for(var i=0;i<arr.length;i++) if(isErr(arr[i])) return arr[i]; return null; }
  function toNum(v){
    if(typeof v==='number') return v;
    if(typeof v==='boolean') return v?1:0;
    if(typeof v==='string' && v.trim()!=='' && !isNaN(v)) return parseFloat(v);
    return null;
  }
  function col(arr){ return SRV.onlyValues(arr.map(function(v){ return [v]; })); }
  function esc(s){ return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  function wildRe(pat){
    var out='';
    for(var i=0;i<pat.length;i++){
      var c=pat.charAt(i);
      if(c==='~' && i+1<pat.length){ out += esc(pat.charAt(i+1)); i++; continue; }
      if(c==='*'){ out += '[\\s\\S]*'; continue; }
      if(c==='?'){ out += '[\\s\\S]'; continue; }
      out += esc(c);
    }
    return new RegExp('^'+out+'$','i');
  }
  function hasWild(s){ return typeof s==='string' && /[*?]/.test(s); }
  function eqVal(a,b){
    if(typeof a==='string' && typeof b==='string') return a.toLowerCase()===b.toLowerCase();
    return a===b;
  }

  //  FunctionPlugin は ES のクラス。ES5の apply では継承できない
  //  (実測: Class constructors cannot be invoked without 'new')
  var ExallyPlugin = class ExallyPlugin extends FunctionPlugin {};

  ExallyPlugin.prototype.exSort = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.SORT'), function(range, idx, order){
      var arr = flat(range).filter(function(v){ return v!==null && v!==undefined && v!==''; });
      var e = firstErr(arr); if(e) return e;
      var dir = (toNum(order)===-1) ? -1 : 1;
      arr.sort(function(a,b){
        if(typeof a==='number' && typeof b==='number') return (a-b)*dir;
        return String(a).localeCompare(String(b))*dir;
      });
      return col(arr);
    });
  };
  ExallyPlugin.prototype.exUnique = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.UNIQUE'), function(range){
      var arr = flat(range).filter(function(v){ return v!==null && v!==undefined && v!==''; });
      var e = firstErr(arr); if(e) return e;
      var seen = [], out = [];
      for(var i=0;i<arr.length;i++){
        var key = (typeof arr[i]) + ' ' + String(arr[i]);
        if(seen.indexOf(key)<0){ seen.push(key); out.push(arr[i]); }
      }
      return col(out);
    });
  };
  ExallyPlugin.prototype.exText = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.TEXT'), function(v, fmt){
      if(isErr(v)) return v;
      if(isErr(fmt)) return fmt;
      var raw = flat(v)[0];
      var n = toNum(raw);
      if(n===null) return raw===null||raw===undefined ? '' : String(raw);
      var r = _applyTextFormat(n, String(fmt));
      return r===null ? String(n) : r;
    });
  };
  ExallyPlugin.prototype.exTextjoin = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.TEXTJOIN'), function(){
      var a = Array.prototype.slice.call(arguments);
      var delim = a[0]===null||a[0]===undefined ? '' : String(flat(a[0])[0]);
      var ign = a[1]===true || String(flat(a[1])[0]).toUpperCase()==='TRUE';
      var vals = [];
      for(var i=2;i<a.length;i++) vals = vals.concat(flat(a[i]));
      var e = firstErr(vals); if(e) return e;
      return _jsTextjoin(delim, ign, vals.map(function(v){ return v===null||v===undefined?'':v; }));
    });
  };
  ExallyPlugin.prototype.exInt = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.INT'), function(v){
      if(isErr(v)) return v;
      var n = toNum(flat(v)[0]);
      if(n===null) return err(ErrorType.VALUE);
      return Math.floor(n);
    });
  };
  ExallyPlugin.prototype.exMod = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.MOD'), function(x, y){
      if(isErr(x)) return x;
      if(isErr(y)) return y;
      var a = toNum(flat(x)[0]), b = toNum(flat(y)[0]);
      if(a===null || b===null) return err(ErrorType.VALUE);
      if(b===0) return err(ErrorType.DIV_BY_ZERO);
      return a - b*Math.floor(a/b);
    });
  };
  ExallyPlugin.prototype.exValue = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.VALUE'), function(v){
      if(isErr(v)) return v;
      var raw = flat(v)[0];
      if(raw===null || raw===undefined || raw==='') return err(ErrorType.VALUE);
      if(typeof raw==='number') return raw;
      var n = _jsValue(raw);
      return n===null ? err(ErrorType.VALUE) : n;
    });
  };
  ExallyPlugin.prototype.exXlookup = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.XLOOKUP'), function(key, la, ra, ifnf){
      if(isErr(key)) return key;
      var k = flat(key)[0];
      var L = flat(la), R = flat(ra);
      var e = firstErr(L) || firstErr(R); if(e) return e;
      for(var i=0;i<L.length;i++){
        if(hasWild(k) ? wildRe(String(k)).test(String(L[i])) : eqVal(L[i], k)) return R[i]===undefined?null:R[i];
      }
      if(ifnf===undefined || ifnf===null) return err(ErrorType.NA);
      return flat(ifnf)[0];
    });
  };
  ExallyPlugin.prototype.exFilter = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.FILTER'), function(range, cond, ifEmpty){
      var A = flat(range), C = flat(cond);
      var e = firstErr(A) || firstErr(C); if(e) return e;
      var out = [];
      for(var i=0;i<A.length;i++){
        var c = C[i];
        if(c===true || c===1 || (typeof c==='string' && c.toUpperCase()==='TRUE')) out.push(A[i]);
      }
      if(!out.length){
        if(ifEmpty===undefined || ifEmpty===null) return err(ErrorType.NA);
        return col(flat(ifEmpty));
      }
      return col(out);
    });
  };
  ExallyPlugin.prototype.exMatch = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.MATCH'), function(key, range, type){
      if(isErr(key)) return key;
      var k = flat(key)[0], A = flat(range);
      var e = firstErr(A); if(e) return e;
      var t = (type===undefined||type===null) ? 1 : toNum(flat(type)[0]);
      if(t===0){
        for(var i=0;i<A.length;i++){
          if(hasWild(k) ? wildRe(String(k)).test(String(A[i])) : eqVal(A[i], k)) return i+1;
        }
        return err(ErrorType.NA);
      }
      var best=-1;
      for(var j=0;j<A.length;j++){
        var v=A[j];
        if(t===1 && typeof v==='number' && typeof k==='number' && v<=k) best=j;
        if(t===-1 && typeof v==='number' && typeof k==='number' && v>=k) best=j;
        if(t===1 && typeof v==='string' && typeof k==='string' && String(v).toLowerCase()<=String(k).toLowerCase()) best=j;
      }
      return best<0 ? err(ErrorType.NA) : best+1;
    });
  };
  ExallyPlugin.prototype.exSumproduct = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.SUMPRODUCT'), function(){
      var a = Array.prototype.slice.call(arguments).map(flat);
      if(!a.length) return 0;
      for(var i=0;i<a.length;i++){ var e=firstErr(a[i]); if(e) return e; }
      var n = a[0].length, sum = 0;
      for(var r=0;r<n;r++){
        var p = 1;
        for(var c=0;c<a.length;c++){
          var v = a[c][r];
          if(v===true) v=1; else if(v===false) v=0;
          var num = typeof v==='number' ? v : ((v===null||v===undefined||v===''||isNaN(v)) ? 0 : parseFloat(v));
          p *= num;
        }
        sum += p;
      }
      return sum;
    });
  };

  // ═══ 第3波P1(2026-08-01) ここから ═══════════════════════════════
  //  _jsSet から移した11関数。入れ子(=ROUND(LOOKUP(...),0) 等)でも正しく計算されるようにする。
  //  真値は tests/xlsx-harness/cases/90-wave3-p1.json ＋ 実Excel(16.0.20228)の golden。
  function txt(v){
    v = unwrap(v);
    if(v===null || v===undefined) return '';
    if(v===true) return 'TRUE';
    if(v===false) return 'FALSE';
    return String(v);
  }
  function optNum(v, dflt){
    if(v===undefined || v===null) return dflt;
    var n = toNum(flat(v)[0]);
    return n===null ? dflt : n;
  }
  function optRaw(v){
    if(v===undefined || v===null) return undefined;
    return unwrap(flat(v)[0]);
  }
  function fromJs(r){ return r==='#N/A' ? err(ErrorType.NA) : (r==='#VALUE!' ? err(ErrorType.VALUE) : (r==='#REF!' ? err(ErrorType.REF) : r)); }

  ExallyPlugin.prototype.exConcat = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.CONCAT'), function(){
      var parts = Array.prototype.slice.call(arguments).map(flat);
      for(var i=0;i<parts.length;i++){ var e=firstErr(parts[i]); if(e) return e; }
      return [].concat.apply([], parts).map(txt).join('');
    });
  };
  ExallyPlugin.prototype.exLookup = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.LOOKUP'), function(key, la, ra){
      if(isErr(key)) return key;
      var k = flat(key)[0], L = flat(la), R = (ra===undefined||ra===null) ? L : flat(ra);
      var e = firstErr(L) || firstErr(R); if(e) return e;
      return fromJs(_jsLookup(k, L, R));
    });
  };
  ExallyPlugin.prototype.exXmatch = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.XMATCH'), function(key, arr, mm, sm){
      if(isErr(key)) return key;
      var k = flat(key)[0], A = flat(arr);
      var e = firstErr(A); if(e) return e;
      return fromJs(_jsXmatch(k, A, optNum(mm, 0), optNum(sm, 1)));
    });
  };
  //  ★INDIRECT は【毎回作り直す(isVolatile)】指定が要る。
  //    プラグインは評価時に値を読むだけで、参照先(E1:E6)は依存グラフに載らない。
  //    volatile を付けないと E3 を書き換えても古い答えが残る＝表示だけ古い最悪の壊れ方になる
  //    (実測: 付けないと 210 のまま／付けると 3180 に更新される)。Excel の INDIRECT も揮発性なので挙動も一致する。
  //  ★対応するのは同じシートの A1形式のみ。'Sheet1!A1' と R1C1形式(第2引数FALSE)は #REF! を返す
  //    ＝黙って違う数字を出さない。台帳 docs/SPEC_wave3_p1_plugin_migration.md に期限つきで記載。
  ExallyPlugin.prototype.exIndirect = function(ast, state){
    var dg = this.dependencyGraph;
    return this.runFunction(ast.args, state, this.metadata('EX.INDIRECT'), function(refStr, a1){
      if(isErr(refStr)) return refStr;
      if(a1!==undefined && a1!==null && flat(a1)[0]===false) return err(ErrorType.REF); // R1C1形式は未対応
      var s = txt(flat(refStr)[0]).trim();
      if(s==='' || s.indexOf('!')>=0) return err(ErrorType.REF);                        // 他シートは未対応
      function one(a){
        var m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(a.trim());
        if(!m) return null;
        var c=0, u=m[1].toUpperCase();
        for(var i=0;i<u.length;i++) c = c*26 + (u.charCodeAt(i)-64);
        return { col:c-1, row:parseInt(m[2],10)-1 };
      }
      var st, en;
      if(s.indexOf(':')>=0){ var p=s.split(':'); st=one(p[0]); en=one(p[1]); }
      else { st=one(s); en=st; }
      if(!st || !en) return err(ErrorType.REF);
      var sheet = state.formulaAddress.sheet;
      var r0=Math.min(st.row,en.row), r1=Math.max(st.row,en.row);
      var c0=Math.min(st.col,en.col), c1=Math.max(st.col,en.col);
      var vals=[];
      for(var r=r0;r<=r1;r++){
        var line=[];
        for(var c=c0;c<=c1;c++) line.push(unwrap(dg.getCellValue({ sheet:sheet, col:c, row:r })));
        vals.push(line);
      }
      if(vals.length===1 && vals[0].length===1) return vals[0][0];
      return SRV.onlyValues(vals);
    });
  };
  ExallyPlugin.prototype.exDatevalue = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.DATEVALUE'), function(s){
      if(isErr(s)) return s;
      var n = _jsDateValue(txt(flat(s)[0]));
      return n===null ? err(ErrorType.VALUE) : n;
    });
  };
  ExallyPlugin.prototype.exNumbervalue = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.NUMBERVALUE'), function(s, d, g){
      if(isErr(s)) return s;
      var n = _jsNumbervalue(txt(flat(s)[0]), optRaw(d), optRaw(g));
      return n===null ? err(ErrorType.VALUE) : n;
    });
  };
  ExallyPlugin.prototype.exFixed = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.FIXED'), function(num, dec, noCommas){
      if(isErr(num)) return num;
      var n = toNum(flat(num)[0]);
      if(n===null) return err(ErrorType.VALUE);
      var nc = optRaw(noCommas);
      return _jsFixed(n, (dec===undefined||dec===null) ? 2 : optNum(dec,2), nc===true || nc===1);
    });
  };
  ExallyPlugin.prototype.exAsc = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.ASC'), function(s){
      if(isErr(s)) return s;
      return _jsAsc(txt(flat(s)[0]));
    });
  };
  ExallyPlugin.prototype.exDbcs = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.DBCS'), function(s){
      if(isErr(s)) return s;
      return _jsDbcs(txt(flat(s)[0]));
    });
  };
  ExallyPlugin.prototype.exTextbefore = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.TEXTBEFORE'), function(t, d, inst, mm, me, inf){
      if(isErr(t)) return t;
      return fromJs(_jsTextbefore(txt(flat(t)[0]), txt(flat(d)[0]), optNum(inst,1), optNum(mm,0), optNum(me,0), optRaw(inf)));
    });
  };
  ExallyPlugin.prototype.exTextafter = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.TEXTAFTER'), function(t, d, inst, mm, me, inf){
      if(isErr(t)) return t;
      return fromJs(_jsTextafter(txt(flat(t)[0]), txt(flat(d)[0]), optNum(inst,1), optNum(mm,0), optNum(me,0), optRaw(inf)));
    });
  };
  // ═══ 第3波P1 ここまで ═══════════════════════════════════════════

  var ANY = { argumentType: T.ANY };
  var OPT = { argumentType: T.ANY, optionalArg: true };
  ExallyPlugin.implementedFunctions = {
    'EX.SORT':       { method: 'exSort',       parameters: [ANY, OPT, OPT, OPT], arrayFunction: true },
    'EX.UNIQUE':     { method: 'exUnique',     parameters: [ANY, OPT, OPT],      arrayFunction: true },
    'EX.TEXT':       { method: 'exText',       parameters: [ANY, ANY] },
    'EX.TEXTJOIN':   { method: 'exTextjoin',   parameters: [ANY, ANY, ANY], repeatLastArgs: 1 },
    'EX.INT':        { method: 'exInt',        parameters: [ANY] },
    'EX.MOD':        { method: 'exMod',        parameters: [ANY, ANY] },
    'EX.VALUE':      { method: 'exValue',      parameters: [ANY] },
    'EX.XLOOKUP':    { method: 'exXlookup',    parameters: [ANY, ANY, ANY, OPT, OPT, OPT] },
    'EX.FILTER':     { method: 'exFilter',     parameters: [ANY, ANY, OPT],      arrayFunction: true },
    'EX.MATCH':      { method: 'exMatch',      parameters: [ANY, ANY, OPT] },
    'EX.SUMPRODUCT': { method: 'exSumproduct', parameters: [ANY], repeatLastArgs: 1 },
    // ★第3波P1(2026-08-01)
    'EX.CONCAT':      { method: 'exConcat',      parameters: [ANY], repeatLastArgs: 1 },
    'EX.LOOKUP':      { method: 'exLookup',      parameters: [ANY, ANY, OPT] },
    'EX.XMATCH':      { method: 'exXmatch',      parameters: [ANY, ANY, OPT, OPT] },
    //  ★isVolatile=毎回作り直す。参照先は依存グラフに載らないので、これが無いと古い答えが残る(実測)。
    'EX.INDIRECT':    { method: 'exIndirect',    parameters: [ANY, OPT], isVolatile: true },
    'EX.DATEVALUE':   { method: 'exDatevalue',   parameters: [ANY] },
    'EX.NUMBERVALUE': { method: 'exNumbervalue', parameters: [ANY, OPT, OPT] },
    'EX.FIXED':       { method: 'exFixed',       parameters: [ANY, OPT, OPT] },
    'EX.ASC':         { method: 'exAsc',         parameters: [ANY] },
    'EX.DBCS':        { method: 'exDbcs',        parameters: [ANY] },
    'EX.TEXTBEFORE':  { method: 'exTextbefore',  parameters: [ANY, ANY, OPT, OPT, OPT, OPT] },
    'EX.TEXTAFTER':   { method: 'exTextafter',   parameters: [ANY, ANY, OPT, OPT, OPT, OPT] }
  };
  var tr = {};
  _PLUGIN_FUNCS.forEach(function(n){ tr['EX.'+n] = n; });
  try {
    P.cls.registerFunctionPlugin(ExallyPlugin, { enGB: tr, enUS: tr });
  } catch(e) {
    if(typeof console!=='undefined') console.warn('Exally関数プラグインの登録に失敗', e);
    return false;
  }
  _pluginRegistered = true;
  return true;
}

// ブラウザ: hyperformula.full.min.js の後に読まれるので、ここで登録する
//  (HyperFormula.buildEmpty より前に登録されている必要がある)
if (typeof HyperFormula !== 'undefined') registerExallyFunctions(HyperFormula);

if (typeof module !== 'undefined' && module.exports) {
  module.exports.registerExallyFunctions = registerExallyFunctions;
  module.exports._PLUGIN_FUNCS = _PLUGIN_FUNCS;
}
