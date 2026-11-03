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
/* ★式のセルが「空のセル」を指した時、Excelは 0 を返す★（2026-08-29 実Excel 16.0 で実測）
 *
 *   =C2（C2は空）         → ★0★（数）  画面に「0」
 *   =INDEX(C1:C5,2)（空）  → ★0★（数）  画面に「0」
 *   =IF(C2="","",1)       → ""（字）    画面は 空
 *   =C2&""                → ""（字）    画面は 空
 *   ⇒ ★決まり＝「数の場所なら 0・字の場所なら 空」★
 *
 *   HyperFormula は 前者で ★null★ を返す（後者は "" を返すので ★区別が つく★）。
 *   ここで null を '' にしていたため、★司さんの実物 19,323本のうち 2,918本が
 *   「Excelは 0・うちは 空」★ で 食い違っていた（Excelとの一致 84.6%）。
 *
 *   ★式のセルだけ★ 0 にする。ただの空セルまで 0 にすると 表が 0 だらけになる。
 *   ⇒ 呼び手が 式かどうかを 渡す（★既定は false＝前と同じ動き★）。
 */
function _hfGetDisplay(sheet, r, c, 式のセルか) {
  try {
    var val = _hf.getCellValue({sheet:_hfSid(sheet), row:r, col:c});
    if(val===null||val===undefined) return 式のセルか ? '0' : '';
    if(typeof val==='object'&&val.type) return HF_ERR[val.type]||('#'+val.type);
    /* ★はい/いいえは 大文字★（実Excel 実測 2026-08-30＝Range.Text が "TRUE"/"FALSE"）。
       String(true) は "true" に なるので ★ここで 直す★。
       ★出口は この 1か所★＝画面も 紙も 書き出しも 同じ 字に なる。 */
    if(typeof val==='boolean') return val ? 'TRUE' : 'FALSE';
    return String(val);
  } catch(e) { return '#ERR'; }
}
/* ★★答えが「文字」か「数」かを 返す★★（2026-09-11）
 *
 *  ★何が 起きて いたか★
 *    `=TEXT(A1,"0.00")` の 答えは 実Excel では ★文字の "3.00"★
 *      実測 … LEN=4 ／ ISTEXT=TRUE ／ 出る字 "3.00"（docs/measured/golden-moji-no-kotae-2026-09-11.tsv）
 *    ★エンジンも 文字で 返して います★（typeof が 'string'）
 *    ★でも `_hfGetDisplay` が `String(val)` に した 時点で ★型が 消える★★
 *    ⇒ 画面が それを ★数として 読み直して★ ★"3" と 出して いました★
 *
 *  ★★TEXT だけの 話では ありません★★
 *    `=LEFT("3.00",4)` `=RIGHT("12.50",5)` `=MID(...)` `=TRIM("  4.50  ")`
 *    `=SUBSTITUTE("1,5",",",".")` `=CONCATENATE(...)` `=REPT("1",3)` `=A1&""`
 *    ⇒★文字を 返す 式は 全部★（実Excel で 11本 確かめた・全部 String）
 *    （`="007"` だけは ★頭の ゼロ★の 決まりで たまたま 守られて いた）
 *
 *  ★返す 物★ … 'string' / 'number' / 'boolean' / 'error' / 'empty'
 *  ★`_hfGetDisplay` と 同じ 物を 見ます★＝ずれない
 */
function _hf答えの型(sheet, r, c) {
  try {
    var val = _hf.getCellValue({sheet:_hfSid(sheet), row:r, col:c});
    if(val===null||val===undefined) return 'empty';
    if(typeof val==='object'&&val.type) return 'error';
    if(typeof val==='boolean') return 'boolean';
    if(typeof val==='string') return 'string';
    if(typeof val==='number') return 'number';
    return 'other';
  } catch(e) { return 'error'; }
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
  //  ★aaa/aaaa(曜日)だけの書式には y/m/d/h/s が1文字も無い。ここで拾わないと
  //    数の書式として扱われ、どの形にも当てはまらず ★シリアル値がそのまま出る★
  //    （司さんの実物 =TEXT(A5,"aaa") 730本が 46023 と表示されていた 2026-08-10 実測）
  //  ★和暦(ggge / gge / ge)も日付の書式。元号の字はまだ訳せない(台帳 TEXT_era_wareki)が、
  //    ここで拾わないと数の書式に落ちて ★シリアル値がそのまま答えになる★。
  //    後ろに英字が続く物(General)は別物なので除く。
  return (/[ymdhs]/i.test(f) || /a{3,4}/i.test(f) || /g{1,3}e(?![a-z])/i.test(f)) && !/[#0]/.test(f);
}

/* ★書式コードで使う名前は 実Excel(16.0.20228・日本語UI 1041)に1つずつ聞いて決めた（2026-08-10 実測）
     aaa=土 / aaaa=土曜日 …日本語。★ddd=Sat / dddd=Saturday / mmm=Jan / mmmm=January / mmmmm=J
     ＝日本語UIのExcelでも英語のまま出る（想像ではなく実測値）。 */
var _WD_JA  = ['日','月','火','水','木','金','土'];
var _WD_EN  = ['Sunday','Monday','Tuesday','Wednesday','Thursday','Friday','Saturday'];
var _MON_EN = ['January','February','March','April','May','June','July','August','September','October','November','December'];

/* 頭に付く言語の指定（[$-411] など）を外す。「その言語で出せ」の指示で、
   何を出すかは a/d/m の側で決まる。実物の書式は m/d\([$-411]aaa\) の形。 */
function _fmtStripLocale(fmt) { return String(fmt).replace(/\[\$-[0-9A-Za-z]+\]/g, ''); }

/* 書式コードを かたまりに切る。引用符の中はまとめて1個（中の m を月にしないため） */
var _FMT_TOK  = /"[^"]*"|am\/pm|a\/p|aaaa|aaa|dddd|ddd|mmmmm|mmmm|mmm|yyyy|yy|mm|m|dd|d|hh|h|ss|s|[\s\S]/gi;
var _FMT_DATE_TOK = /^(yyyy|yy|mmmmm|mmmm|mmm|mm|m|dddd|ddd|dd|d|aaaa|aaa|hh|h|ss|s)$/i;
/* ★m は「月」か「分」か★ = 時(h)のすぐ後、または秒(s)のすぐ前なら分。
   実Excel実測: TEXT(46053.5,"hh:mm")="12:00"（月なら 12:01 になる）／ "mm:ss"="00:00" */
function _fmtIsMinute(toks, i) {
  var j, t;
  for (j=i-1; j>=0; j--) { t=toks[j]; if(t.charAt(0)==='"') continue; if(_FMT_DATE_TOK.test(t)) return /^hh?$/i.test(t); }
  for (j=i+1; j<toks.length; j++) { t=toks[j]; if(t.charAt(0)==='"') continue; if(_FMT_DATE_TOK.test(t)) return /^ss?$/i.test(t); }
  return false;
}
// シリアル値 → 日付(1900系。Excelの1900年閏年バグ域=シリアル60以前は扱わない)
function _fmtDate(serial, fmt) {
  var days = Math.floor(serial);
  var ms   = Math.round((serial - days) * 86400) * 1000;
  var d    = new Date(Date.UTC(1899,11,30) + days*86400000 + ms);
  var Y=d.getUTCFullYear(), Mo=d.getUTCMonth()+1, D=d.getUTCDate(), W=d.getUTCDay();
  var H=d.getUTCHours(), Mi=d.getUTCMinutes(), S=d.getUTCSeconds();
  var p2=function(v){ return (v<10?'0':'')+v; };
  var toks = String(fmt).match(_FMT_TOK) || [];
  var ampm = toks.some(function(t){ return /^(am\/pm|a\/p)$/i.test(t); });
  var Hd   = ampm ? (H%12===0 ? 12 : H%12) : H;   // 実測 TEXT(0.75,"h:mm AM/PM")="6:00 PM"
  var out = '';
  for (var i=0;i<toks.length;i++) {
    var t = toks[i];
    if (t.charAt(0)==='"') { out += t.slice(1,-1); continue; }
    switch (t.toLowerCase()) {
      case 'yyyy':  out += String(Y); break;
      case 'yy':    out += p2(Y%100); break;
      case 'mmmmm': out += _MON_EN[Mo-1].charAt(0); break;
      case 'mmmm':  out += _MON_EN[Mo-1]; break;
      case 'mmm':   out += _MON_EN[Mo-1].slice(0,3); break;
      case 'mm':    out += _fmtIsMinute(toks,i) ? p2(Mi) : p2(Mo); break;
      case 'm':     out += _fmtIsMinute(toks,i) ? String(Mi) : String(Mo); break;
      case 'dddd':  out += _WD_EN[W]; break;
      case 'ddd':   out += _WD_EN[W].slice(0,3); break;
      case 'dd':    out += p2(D); break;
      case 'd':     out += String(D); break;
      case 'aaaa':  out += _WD_JA[W]+'曜日'; break;
      case 'aaa':   out += _WD_JA[W]; break;
      case 'hh':    out += p2(Hd); break;
      case 'h':     out += String(Hd); break;
      case 'ss':    out += p2(S); break;
      case 's':     out += String(S); break;
      case 'am/pm': out += (H<12?'AM':'PM'); break;
      case 'a/p':   out += (H<12?'A':'P'); break;
      default:      out += t;
    }
  }
  return out;
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
/* ══ ★★和暦（g・e）と 経過時間（[h] [m] [s]）は 書式の 台（lib/shoshiki.js）に 渡す★★ ══（2026-10-02）
     ★この 層の 書式器は 和暦・経過時間を 知りません★（実測 `[h]:mm` → [12]:00 ／ `ge.m.d` → ge.1.31）
     ★台は 実Excel 16.0.20430 の 真値 155本と 全部 一致★（tests/xlsx-harness/cases/61-text-wareki-keika.json）
     ⇒★台に 渡します★＝★2か所で 解かない★（画面と 自前の 台の TEXT も 台を 呼ぶ）
     ★返り★ undefined＝和暦でも 経過でも ない（今まで通り）／null＝★出せない（#VALUE!）★／字＝答え
     ★判じは 台の 1つだけ★（2026-10-04・画面の fmtForDisplay も 同じ 判じを 呼ぶ）
       ★前は ここに 自前の 判じが 在り `0.00E+00` の 指数の E を 和暦と 取り違えて いた★
       （⇒ 台へ 渡って #VALUE!。台の 判じは 指数の E・General・引用符・逃がし・[Red] 等を 外す）
     ★★日付の 最後の 日（通し 2958465）から 先も 台へ★★（10-04 夕・画面の fmtForDisplay と 同じ 決まり）
       この 層の _fmtDate は 上限を 知らず `=TEXT(2958466,"yyyy/m/d")` を ★10000/1/1★ に して いた（実Excel #VALUE!） */
function _和暦か経過なら台へ(num, fmt) {
  if (typeof Shoshiki === 'undefined' || !Shoshiki || typeof Shoshiki.字にする !== 'function') return undefined;
  var 和経 = typeof Shoshiki.和暦か経過か === 'function' && Shoshiki.和暦か経過か(String(fmt));
  var 上限 = typeof num === 'number' && num >= 2958465
    && typeof Shoshiki.日付の書式か === 'function' && Shoshiki.日付の書式か(String(fmt));
  if (!和経 && !上限) return undefined;
  return Shoshiki.字にする(num, String(fmt));
}
function _applyTextFormat(num, fmt) {
  var f = _fmtStripLocale(fmt);
  //  ★日付の書式は「正;負;ゼロ;文字」の1区画目だけを使う（実測 TEXT(46053,"aaa;@")="土"）
  if(_fmtIsDate(f)) return _fmtDate(num, _fmtSections(f)[0]);
  return _fmtNumber(num, f);
}

// --- 統計 ---
// RANK / RANK.EQ / RANK.AVG。order は 0=降順(既定) / 0以外=昇順。範囲に無ければ #N/A。
//  EQ=同順位は同じ番号(小さい方) / AVG=同順位は平均した番号。
function _jsRank(val,vals,order,avg){
  var nums=vals.filter(function(v){return typeof v==='number'&&isFinite(v);});
  var s=nums.slice().sort(function(a,b){return (order===undefined||order===null||Math.trunc(order)===0)?b-a:a-b;});
  var first=s.indexOf(val);
  if(first<0) return '#N/A';
  if(!avg) return first+1;
  var last=s.lastIndexOf(val);
  return (first+1+last+1)/2;
}
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
/* ★★`_一本の並びか()` は 外しました（2026-09-09）★★
   「x が 2本以上なら 断る」為に 作った 門でした。
   ⇒★断るのを やめて 直した★ので ★誰からも 呼ばれません★
   ⇒★呼ばれない 物を 置くと「そこで 見て いる」と 読まれる★＝消します
   （今は `lib/formula-yosoku.js` の `直線の係数()` が x を 何本でも 解きます） */
/* ★★LINEST／LOGEST／TREND／GROWTH の 下請け 4本は 外しました（2026-09-11）★★
   ★誰からも 呼ばれて いませんでした★（数えた … 呼び出し 0本／試験 0本／道具 0本）
     `_jsGrowth` が `_jsLogest` を 呼ぶ 1本だけが 4本の 中の 行き来でした。
   ★今は `lib/formula-yosoku.js` の `直線の係数()` が ★表として★ 返します★
     ・1つの 数しか 返せない 段では LINEST は 出せません（実測 … 実Excel 0.7708… ／ うち 7.390…）
     ・実測 … docs/measured/golden-linest-hyou-2026-09-09.tsv（実Excel 119本）
   ⇒★呼ばれない 物を 置くと「そこで 見て いる」と 読まれます★（上の `_一本の並びか()` と 同じ）
   ★司さん（2026-09-02）「使わんもんは 消せや」★ */

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
// ── DOLLAR（通貨書式）─────────────────────────────────────────────
//  ★通貨書式の関数の【本名は DOLLAR】。YEN は日本語UIの表示名で、Excel自身が
//    .FormulaLocal(=YEN(...)) ↔ .Formula(=DOLLAR(...)) で相互変換する(実測)。
//    =YEN(...) をファイル/US-English構文で書くと #NAME? になるので、
//    入口(convertFormula)と出口(xlsx書き出し)で DOLLAR に寄せる。
//  ★出す通貨記号と既定の桁数は【地域の通貨書式】に従う。このハーネスの前提環境
//    (日本語1041 / 国別81)での実測は ¥ ・小数0桁 ・負は「¥-1,235」(括弧ではない)。
//    RECIPE.md に書いてあるロケール前提と同じ扱い＝環境が変わればここも変わる。
var _CURRENCY = { symbol: '¥', decimals: 0 };
function _jsDollar(num,dec){
  dec = (dec===undefined||dec===null) ? _CURRENCY.decimals : Math.trunc(dec);
  var n = _xlRound(num, dec);
  var neg = n < 0;
  var str = Math.abs(n).toFixed(Math.max(0,dec));
  var p = str.split('.');
  p[0] = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return _CURRENCY.symbol + (neg ? '-' : '') + p.join('.');
}
var _jsYen = _jsDollar;   // 旧名(日本語UI名)。中身は同じ物を指す＝二重実装にしない。
// ── バイト単位の文字列関数(LENB/LEFTB/RIGHTB/MIDB) ────────────────────
//  ★日本語Excelのバイト数(Shift-JIS基準): ASCII=1 / ★半角カナ=1 / それ以外(全角)=2。
//    実測: LENB("ｱｲｳ")=3(半角カナは1バイト)、LENB("りんご")=6、LENB("山田 太郎")=9。
//  ★2バイト文字の途中で切ると、その【半分は空白1個】になる(消えるのではない)。実測:
//    LEFTB("りんご",3)="り " / RIGHTB("りんご",3)=" ご" / MIDB("りんご",2,2)="  "(空白2個)
function _byteWidth(c){
  var cc=c.charCodeAt(0);
  if(cc<=0x7F) return 1;
  if(cc>=0xFF61 && cc<=0xFF9F) return 1;   // 半角カナ・半角句読点
  return 2;
}
function _jsLenb(s){var n=0;for(var i=0;i<s.length;i++)n+=_byteWidth(s.charAt(i));return n;}
// 1始まりのバイト位置から numBytes バイトを取り出す。半分だけ入った2バイト文字は空白になる。
function _byteSlice(s,startByte,numBytes){
  if(numBytes<=0) return '';
  var units=[];   // 1バイトごとに「どの文字の何番目の半分か」を並べる
  for(var i=0;i<s.length;i++){
    var c=s.charAt(i), w=_byteWidth(c);
    if(w===1) units.push({c:c,whole:true});
    else { units.push({c:c,idx:i,half:1}); units.push({c:c,idx:i,half:2}); }
  }
  var from=Math.max(1,Math.trunc(startByte))-1;
  var slice=units.slice(from, from+Math.trunc(numBytes));
  var out='';
  for(var k=0;k<slice.length;k++){
    var u=slice[k];
    if(u.whole){ out+=u.c; continue; }
    if(u.half===1 && k+1<slice.length && slice[k+1].half===2 && slice[k+1].idx===u.idx){ out+=u.c; k++; continue; }
    out+=' ';   // 片方の半分しか入っていない=空白1個
  }
  return out;
}
function _jsLeftb(s,b){return _byteSlice(s,1,b);}
function _jsRightb(s,b){var n=_jsLenb(s);return _byteSlice(s,n-Math.trunc(b)+1,b);}
function _jsMidb(s,start,len){return _byteSlice(s,start,len);}
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
// VALUETOTEXT(値,[書式])。書式 0=簡潔(既定・文字列に引用符を付けない) / 1=厳密(文字列を引用符で囲む)
function _jsValuetotext(v,format){
  var strict = Math.trunc(format||0)===1;
  if(typeof v==='string') return strict ? '"'+v+'"' : v;
  if(typeof v==='boolean') return v?'TRUE':'FALSE';
  if(v===null||v===undefined) return '';
  return String(v);
}
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
/* ★★XIRR の 組を 作る（★門と 計算が 同じ 物を 見る 為★）★★（2026-09-09）
 *   ★門と 計算で 別々に 組を 作ると 口裏が 合いません★
 *     ⇒ 実際 踏みました … 門は 通したのに 計算側が `filter(typeof v==='number')` で
 *       ★空の マスを 落として 日付と ずれた★
 *       AA1=-1000／AA2=★空★／AA3=1100 … 実Excel 0.21063382029533387 ／ うち 0.46563424984948343
 *   ★空の マスは 0★（実測）／★日付が 無い 組は 数えない★
 *   @returns {{組:Array,値:Array,日:Array}}
 */
function _xirrの組(生値,生日){
  var 組=[],値=[],日=[];
  for(var i=0;i<生値.length;i++){
    var v=生値[i],d=生日[i];
    var 値が空=(v===null||v===undefined||v==='');
    var 日が空=(d===null||d===undefined||d==='');
    if(日が空)continue;
    var a=値が空?0:Number(v),b=Number(d);
    組.push([a,b]); 値.push(a); 日.push(b);
  }
  return {組:組,値:値,日:日};
}

/* ★★XIRR の 門★★（2026-09-09）
 *   ★どこから 断るかは 私が 決めて いません＝実Excel に 打たせました★
 *   実測 … `docs/measured/golden-xirr-sakaime-2026-09-09.tsv`（61本）
 *
 *   ★実Excel が 断る 条件★
 *     ・字が 混ざる ………………………… ★#VALUE!★
 *     ・値と 日付の 数が 違う …………… ★#NUM!★
 *     ・お金が 1件だけ …………………… ★#N/A★
 *     ・符号が 片方だけ（全部＋／全部−／全部0）… ★#NUM!★
 *     ・日付が 昇順で ない（逆順・バラバラ）…… ★#NUM!★
 *     ・日付に 負の 数 …………………… ★#NUM!★
 *   ★実Excel が 断らない 条件（★止めては いけない★）★
 *     ・空の マスが 混ざる ……… 値を 返す（0.21063382029533387）
 *     ・同じ 日付が 2つ ………… 値を 返す（2.071931099891663）
 *     ・0 を 含む ………………… 値を 返す（0.21063382029533387）
 *
 *   ★見当を 変えても 誤りは 変わりません★（0.1 / 0 / -0.9 / 10 の 4通りで 確かめた）
 *   ⇒★繰り返し 計算の 失敗では なく ★入口の 門★★
 *
 *   @returns {string|null} 誤りの 字（'#NUM!' など）／断らないなら null
 */
function _xirrの門(生値,生日,見当){
  var i;
  /* ★字が 混ざる＝#VALUE!★（★空は 字では ない★＝下で 落とす） */
  for(i=0;i<生値.length;i++){
    var a=生値[i];
    if(typeof a==='string'&&a!==''){ if(!isFinite(Number(a)))return '#VALUE!'; }
    if(typeof a==='boolean')return '#VALUE!';
  }
  for(i=0;i<生日.length;i++){
    var b=生日[i];
    if(typeof b==='string'&&b!==''){ if(!isFinite(Number(b)))return '#VALUE!'; }
    if(typeof b==='boolean')return '#VALUE!';
  }
  /* ★★範囲の 大きさが 違う＝#NUM!★★（★中身の 数では ない★）
     実測 … `=XIRR(X1:X3,Y1:Y2)`（3マス と 2マス）… ★#NUM!★ */
  if(生値.length!==生日.length)return '#NUM!';
  /* ★★空の マスは ★0★ として 扱う（★落とさない★）★★
     実測 … AA1=-1000／AA2=★空★／AA3=1100 と I1=-1000／I2=0／I3=1100 が
            ★どちらも 0.21063382029533387★＝★空は 0★
     ⇒★1回目 私は「空は 落とす」と 書き、★数が 違う→#NUM!★ に して いました★
       ⇒ 実Excel は 値を 返すのに ★うちだけ 断る★＝★出来る 物を 止める★所でした */
  var 組=_xirrの組(生値,生日).組;
  /* ★1件だけ＝#N/A★（★#NUM! では ない★＝実測） */
  if(組.length<2)return '#N/A';
  /* ★日付に 負の 数＝#NUM!★ */
  for(i=0;i<組.length;i++)if(!(組[i][1]>=0))return '#NUM!';
  /* ★日付が 昇順で ない＝#NUM!★（★同じ 日付は 通る★＝実測） */
  for(i=1;i<組.length;i++)if(組[i][1]<組[i-1][1])return '#NUM!';
  /* ★符号が 片方だけ＝#NUM!★（★プラスと マイナスが 両方 要る★） */
  var 正が在る=false,負が在る=false;
  for(i=0;i<組.length;i++){ if(組[i][0]>0)正が在る=true; if(組[i][0]<0)負が在る=true; }
  if(!正が在る||!負が在る)return '#NUM!';
  /* ★★見当が 0 より 小さい＝#NUM!★★（実測）
     ★1回目 私は「−1 以下」と 決めました★（−1／−2／−0.9 の 3つしか 打って いなかった）
     ⇒★刻んで 打ち直したら ★−0.01 でも #NUM!★＝私の 決めが 間違い★
     実測（`=XIRR(E1:E3,F1:F3,見当)`）
       −0.99 −0.95 −0.9 −0.8 −0.7 −0.5 −0.3 −0.1 ★−0.01★ … 全部 ★#NUM!★
       ★0★ … 値（1.0011933300781244）／0.01 0.5 1 2 5 100 10000 … 値
     ⇒★境目は 0★（0 は 通る・0 より 小さいと 断る）
     ⇒★境目は 私が 決めず ★刻んで 打ちました★★ */
  if(見当!==undefined){
    if(!isFinite(見当))return '#VALUE!';
    if(見当<0)return '#NUM!';
  }
  return null;
}
/* ★★見当が 大きいと 発散して NaN に なって いた（2026-09-09 に 直した）★★
 *   =XIRR(E1:E3,F1:F3,10)     … 実Excel 1.0011933278292418 ／ うち ★NaN★
 *   =XIRR(G1:G2,H1:H2,10)     … 実Excel 0.1000000024214387 ／ うち ★-2038322207737022.2★
 *   =XIRR(E1:E3,F1:F3,1000000)… 実Excel 1.0011933255782424 ／ うち ★NaN★
 *   ⇒ ニュートン法が ★1+rate が 0 か 負★の 所へ 飛び、`Math.pow` が NaN に なる
 *   ⇒★実Excel は どの 見当でも 同じ 所に 落ち着きます★（実測 … 0.01〜10000 で 全部 1.00119…）
 *   ★直し方★
 *     ①★1+rate が 0 以下に なったら 戻す★（−1 より 下へ 行かせない）
 *     ②★NaN に なったら 打ち切って 見当を 0.1 から やり直す★
 *     ⇒★★出来る 物を 止めない★★（断るのでは なく 落ち着かせる）
 */
function _jsXirr(values,dates,guess){
  var 始め=(guess===undefined)?0.1:guess;
  var r=_xirrを解く(values,dates,始め);
  /* ★落ち着かなければ 0.1 から やり直す★（実Excel は 見当が 何でも 同じ 所に 落ちる） */
  if(!isFinite(r)&&始め!==0.1)r=_xirrを解く(values,dates,0.1);
  return r;
}
function _xirrを解く(values,dates,guess){
  var d0=dates[0],rate=guess;
  for(var iter=0;iter<1000;iter++){
    var npv=0,dnpv=0,壊れた=false;
    for(var i=0;i<values.length;i++){
      var t=(dates[i]-d0)/365;
      var 底=1+rate;
      if(!(底>0)){壊れた=true;break;}      /* ★1+rate が 0 以下＝累乗が 成り立たない★ */
      var f=Math.pow(底,t);
      if(!isFinite(f)||f===0){壊れた=true;break;}
      npv+=values[i]/f;
      dnpv-=t*values[i]/Math.pow(底,t+1);
    }
    if(壊れた)return NaN;
    if(!isFinite(npv)||!isFinite(dnpv))return NaN;
    if(Math.abs(dnpv)<1e-15)break;
    var nr=rate-npv/dnpv;
    if(!isFinite(nr))return NaN;
    /* ★★−1 より 下へ 行かせない★★（行くと 次の 回で 必ず 壊れる）
       ⇒ 半分だけ 戻す＝★ゆっくり 近づける★ */
    if(nr<=-1)nr=(rate-1)/2;
    if(Math.abs(nr-rate)<1e-10){rate=nr;break;}
    rate=nr;
  }
  return rate;
}
function _jsVdb(cost,salvage,life,start,end,factor){factor=factor||2;var total=0;for(var p=Math.floor(start);p<Math.ceil(end);p++){var s=Math.max(start,p),e2=Math.min(end,p+1);var book=cost;for(var i=0;i<p;i++)book-=Math.min(book*factor/life,book-salvage);total+=Math.min(book*factor/life,book-salvage)*(e2-s);}return total;}
function _jsDisc(settlement,maturity,pr,redemption){var t=(maturity-settlement)/365;return(redemption-pr)/(redemption*t);}
function _jsIntrate(settlement,maturity,investment,redemption){var t=(maturity-settlement)/365;return(redemption-investment)/(investment*t);}
function _jsDuration(settlement,maturity,coupon,yld,freq){var n=Math.round((maturity-settlement)/365*freq);var pv=0,wsum=0;for(var k=1;k<=n;k++){var cf=coupon*100/freq+(k===n?100:0);var pvk=cf/Math.pow(1+yld/freq,k);pv+=pvk;wsum+=pvk*k/freq;}return wsum/pv;}
function _jsAccrint(issue,firstInterest,settlement,rate,par,freq){var t=(settlement-issue)/365;return par*rate*t;}
function _jsAmordegrc(cost,date_purchased,first_period,salvage,period,rate){return cost*rate;}
function _jsAmorlinc(cost,date_purchased,first_period,salvage,period,rate){return(cost-salvage)*rate;}

// --- エンジニアリング ---
/* ★_jsConvert は 2026-09-07 に 消した★
   ⇒ 中の 表が ★手で 書いた 概数★だった（lbm 453.592／実Excel 453.59237、
     cal 4.184／実Excel は 4.1868・4.184 は "c"、gal 0.003785／実Excel 0.003785411784）
   ⇒★どこからも 呼ばれていない事を 確かめてから 消した★
   ⇒ 今は lib/tanni-hyou.js（★実Excel から 機械で 作った 表★）を 使う */
function _jsGestep(num,step){return num>=(step||0)?1:0;}

// --- データベース ---
/* ★条件の マス 1つに 実の 値が 当てはまるか★（2026-09-15）
   ★実 Excel の 実測★（tests/dkei6-honban.test.mjs の 26行）
     `>3`  … 4,5 ⇒ ★9★ ／ `<>2` … 3,4,5 ⇒ ★12★ ／ `>=3` … 3,4,5 ⇒ ★12★
   ★前は 字どうしの `===` だけ★＝`'>3'` が どの マスとも 合わず ★0★ を 返して いた
   ★空の 条件は 縛らない★（実測 … 条件の マスが 空 ⇒ ★14★＝全部 足す） */
function _jsDbAtehamaru(v,c){
  if(c===null||c===undefined||c==='')return true;
  if(typeof c==='string'){
    var m=/^(>=|<=|<>|>|<|=)([\s\S]*)$/.exec(c);
    if(m){
      var r=m[2].trim(),rn=Number(r),vn=Number(v);
      var 数どうし=(r!==''&&isFinite(rn)&&typeof v==='number');
      var d=数どうし?(vn-rn):String(v).localeCompare(r);
      if(m[1]==='>')return d>0;
      if(m[1]==='<')return d<0;
      if(m[1]==='>=')return d>=0;
      if(m[1]==='<=')return d<=0;
      if(m[1]==='<>')return d!==0;
      return d===0;
    }
  }
  return String(v)===String(c);
}
function _jsDbFunc(func,sheet,dbRange,field,criteriaRange){if(!_hf)return'#VALUE!';var m=dbRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);if(!m)return'#VALUE!';var s=_toRC(m[1]),e=_toRC(m[2]);var headers=[];for(var c=s.c;c<=e.c;c++)headers.push(_hf.getCellValue({sheet:_hfSid(sheet),row:s.r,col:c}));var fieldIdx=typeof field==='number'?field-1:headers.indexOf(field);/* ★field が 台帳の 列数より 大きい／0／負★ … 実 Excel は ★#VALUE!★ ★前は 門が `<0` だけで、5 を 渡すと ★台帳の 外の 列★を 読んで ★0★ を 返して いた ★実測★ `=DSUM(A1:B5,5,F1:F2)` … 前 ★0★ ／実 Excel ★#VALUE!★（tests/dkei6-honban.test.mjs の 26行） */if(fieldIdx<0||fieldIdx>=headers.length)return'#VALUE!';var cm=criteriaRange.match(/^([A-Z]+\d+):([A-Z]+\d+)$/i);if(!cm)return'#VALUE!';var cs=_toRC(cm[1]),ce=_toRC(cm[2]);/* ★条件は ★1列とは 限りません★★（2026-09-15・実 Excel の 実測） ★横に 並んだ 列は ★かつ★★ … `F1=1 G1=2 / F2=2 G2=4` ⇒ A=2 ★かつ★ B=4 ⇒ ★2★ `F2=2 G2=99` … 片方 外れ ⇒ ★0★（★前は 1列目しか 見ず 2 を 返して いた★） ★見出しが 台帳に 無い 列は 縛らない★（`G1=99` … ⇒ ★2★） */var 条列=[];for(var cc=cs.c;cc<=ce.c;cc++){var h=_hf.getCellValue({sheet:_hfSid(sheet),row:cs.r,col:cc});条列.push({先:headers.indexOf(h),列:cc});}/* ★★D系の 条件の 決め（2026-09-15・実 Excel 26行で 測って 直した）★★
     紙 `docs/measured/kansuu46/golden-dkei6-kane3-2026-09-15.tsv`（16.0 build 20326）
     見張り `tests/dkei6-honban.test.mjs`（★27 passed / 0 failed★）
   ①★横に 並んだ 列は ★かつ★★（前は 1列目しか 見て いなかった）
   ②★縦に 並んだ 行は ★または★★（前は 2行目しか 見て いなかった）
   ③★見出しが 台帳に 無い 列は ★縛らない★★（前は -1 を 足して 範囲の 外を 読んで いた）
   ④★空の マスも 縛らない★
   ⑤★条件の 行が 無ければ 全部 採る★
   ★台（`lib/shiki-kansuu.js` の `D系の値たち`）と ★同じ 向き★に 揃えて あります★
     ＝★作る道が 2本 在る時は 両方 直せ★ */var results=[];for(var r=s.r+1;r<=e.r;r++){/* ★縦に 並んだ 行は ★または★★（2026-09-15・実 Excel の 実測） `F1=1 / F2=2 / F3=3` … A=2 ★または★ A=3 ⇒ ★5★ ★前は `cs.r+1`（2行目）しか 見ず ★2★ を 返して いた★ ★条件の 行が 無ければ 全部 採る★（＝見出しだけ の 四角） */var 合うか=(ce.r<=cs.r);for(var cr=cs.r+1;cr<=ce.r&&!合うか;cr++){var 全部=true;for(var k=0;k<条列.length&&全部;k++){if(条列[k].先<0)continue;var 条値=_hf.getCellValue({sheet:_hfSid(sheet),row:cr,col:条列[k].列});if(条値===null||条値===undefined||条値==='')continue;if(!_jsDbAtehamaru(_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:s.c+条列[k].先}),条値))全部=false;}if(全部)合うか=true;}if(合うか){results.push(_hf.getCellValue({sheet:_hfSid(sheet),row:r,col:s.c+fieldIdx}));}}var nums=results.filter(function(v){return typeof v==='number';});if(func==='SUM')return nums.reduce(function(a,b){return a+b;},0);if(func==='AVG')return nums.length?nums.reduce(function(a,b){return a+b;})/nums.length:'#DIV/0!';if(func==='CNT')return nums.length;if(func==='CNTA')return results.filter(function(v){return v!==null&&v!==undefined;}).length;if(func==='MAX')return nums.length?Math.max.apply(null,nums):'#NUM!';if(func==='MIN')return nums.length?Math.min.apply(null,nums):'#NUM!';/* ★DPRODUCT は 0件の 時 ★0★★（実 Excel・実測） ★前は `reduce` の 元が 1 なので ★1★ を 返して いた★ ＝`=DPRODUCT(A1:B5,1,F1:F2)` 0件 … 前 ★1★ ／実 Excel ★0★ ★掛け算の 元は 1 なのに 0 を 返す★＝★当て推量では 出ない 形★（PRODUCT も 同じ） */if(func==='PROD')return nums.length?nums.reduce(function(a,b){return a*b;},1):0;/* ★DGET は 0件と 2件以上で ★誤りが 違います★（実 Excel・実測） 0件 … ★#VALUE!★ ／ 2件以上 … ★#NUM!★ ★前は どちらも #NUM!★（`=DGET(A1:B5,1,F1:F2)` 0件 … 前 #NUM! ／実 Excel #VALUE!） */if(func==='GET')return results.length===1?results[0]:(results.length===0?'#VALUE!':'#NUM!');if(func==='STD'){var n=nums.length;if(n<2)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return Math.sqrt(nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/(n-1));}if(func==='STDP'){var n=nums.length;if(!n)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return Math.sqrt(nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/n);}if(func==='VAR'){var n=nums.length;if(n<2)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/(n-1);}if(func==='VARP'){var n=nums.length;if(!n)return'#DIV/0!';var mean=nums.reduce(function(a,b){return a+b;})/n;return nums.reduce(function(a,v){return a+Math.pow(v-mean,2);},0)/n;}return'#VALUE!';}

// --- Web ---
// ENCODEURL。★JSの encodeURIComponent は ! ' ( ) * ~ を素通しするが、Excelはこれらも変換する。
//   実測: =ENCODEURL("!'()*-_.~") → %21%27%28%29%2A-_.%7E （残るのは A-Z a-z 0-9 - _ . だけ）
function _jsEncodeUrl(str){
  try{
    return encodeURIComponent(str).replace(/[!'()*~]/g, function(c){
      return '%' + c.charCodeAt(0).toString(16).toUpperCase();
    });
  }catch(e){ return '#VALUE!'; }
}

// --- AGGREGATE ---
function _jsAggregate(funcNum,opts,vals,k){
  // ★集計番号は Excel の定義に従う(実測で確認): 1 AVERAGE / 2 COUNT / 3 COUNTA / 4 MAX / 5 MIN /
  //   6 PRODUCT / 7 STDEV.S / 8 STDEV.P / 9 SUM / 10 VAR.S / 11 VAR.P /
  //   ★12 MEDIAN / ★13 MODE.SNGL（この2つは以前の実装で逆になっていた）/
  //   14 LARGE / 15 SMALL / 16 PERCENTILE.INC / 17 QUARTILE.INC / 18 PERCENTILE.EXC / 19 QUARTILE.EXC
  //   14〜19 は第4引数 k を取る。対応していない番号は #VALUE! を返す＝黙って違う数字を出さない。
  var clean=vals.filter(function(v){return typeof v==='number'&&isFinite(v);});
  var sum=function(a){return a.reduce(function(x,y){return x+y;},0);};
  var mean=function(a){return sum(a)/a.length;};
  var sorted=function(){return clean.slice().sort(function(a,b){return a-b;});};
  var n=clean.length;
  switch(Math.trunc(funcNum)){
    case 1: return n? mean(clean) : '#DIV/0!';
    case 2: return n;
    case 3: return vals.filter(function(v){return v!==null&&v!==undefined&&v!=='';}).length;
    case 4: return n? Math.max.apply(null,clean) : 0;
    case 5: return n? Math.min.apply(null,clean) : 0;
    case 6: return clean.reduce(function(a,b){return a*b;},1);
    case 7: { if(n<2) return '#DIV/0!'; var m=mean(clean); return Math.sqrt(sum(clean.map(function(v){return Math.pow(v-m,2);}))/(n-1)); }
    case 8: { if(!n) return '#DIV/0!'; var m8=mean(clean); return Math.sqrt(sum(clean.map(function(v){return Math.pow(v-m8,2);}))/n); }
    case 9: return sum(clean);
    case 10:{ if(n<2) return '#DIV/0!'; var m10=mean(clean); return sum(clean.map(function(v){return Math.pow(v-m10,2);}))/(n-1); }
    case 11:{ if(!n) return '#DIV/0!'; var m11=mean(clean); return sum(clean.map(function(v){return Math.pow(v-m11,2);}))/n; }
    case 12:{ var s12=sorted(); if(!s12.length) return '#NUM!'; return s12.length%2 ? s12[(s12.length-1)/2] : (s12[s12.length/2-1]+s12[s12.length/2])/2; }
    case 13: return _jsMode(clean);
    case 14:{ var s14=sorted(); var kk=Math.trunc(k); if(!kk||kk<1||kk>s14.length) return '#NUM!'; return s14[s14.length-kk]; }
    case 15:{ var s15=sorted(); var k15=Math.trunc(k); if(!k15||k15<1||k15>s15.length) return '#NUM!'; return s15[k15-1]; }
    case 16: return (k===undefined||k===null)? '#NUM!' : _jsPercentile(clean,k);
    case 17: return (k===undefined||k===null)? '#NUM!' : _jsQuartile(clean,Math.trunc(k));
    default: return '#VALUE!';   // 18/19(EXC系)は未対応＝黙って近い値を出さない
  }
}

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

/* ★★かっこの 外の カンマで 割る★★（2026-09-18）
     ＝`LAMBDA(x,y,x)` を `['x','y','x']` に する
     ＝★かっこの 中の カンマ（`SUM(1,2)`）では 割りません★ */
function _topLevelSplit(s) {
  var 出=[],深=0,頭=0,中の字=null;
  for(var i=0;i<s.length;i++){
    var c=s.charAt(i);
    if(中の字){ if(c===中の字) 中の字=null; continue; }
    if(c==='"'||c==="'"){ 中の字=c; continue; }
    if(c==='('||c==='['||c==='{') 深++;
    else if(c===')'||c===']'||c==='}') 深--;
    else if(c===','&&深===0){ 出.push(s.slice(頭,i)); 頭=i+1; }
  }
  出.push(s.slice(頭));
  return 出;
}

// MAP(range, LAMBDA(x, body)) → 全件カンマ区切り
function _jsMapCompute(sheet, rangeStr, lambdaFormula) {
  if(!_hf) return null;
  /* ★★引数の 数が 違う LAMBDA を 飲み込んで いました★★（2026-09-18・★お客さんに 出て いた★）
       `=MAP(A1:A5,LAMBDA(x,y,x))`
         ★実Excel★ … `#VALUE!`（実測・`golden-kansuu-9kaime-2026-09-18.tsv` 55行目）
         ★前の うち★ … `#ERROR!`（★実Excel に 無い 字★）
       ★因★ … 下の ふるいが `([^,)]+)` で ★口を 1つだけ★ 取り、
              残り `y,x` を ★そのまま 体★に して いました
              ⇒ `=y,(1)` の ような 式を 借り物に 渡し ⇒ `#ERROR!`
       ★直し★ … ★口の 数を 数え、1つで なければ JS層は 答えません★（`null`）
              ⇒★台（`lib/shiki-hyou.js`）が 答えます★＝`#VALUE!`（実Excel と 同じ）
              ⇒★半分 合う 答えを 出さない★
       ★`#ERROR!`（`!` 有り）は JS層から／`#ERROR`（`!` 無し）は 借り物から★
         ＝★字が 1文字 違うだけで 出どころが 違います★（経営者1 の 見つけ物） */
  var 中=lambdaFormula.replace(/^LAMBDA\s*\(/i,'').replace(/\)\s*$/,'');
  var 割=_topLevelSplit(中);
  if(割.length!==2) return null;   /* ★口1つ ＋ 体1つ＝2つ以外は 答えません★ */
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

/* ★★`_jsIsomitted` は 2026-09-08 に 消した★★
   ⇒ 中身は `v===undefined||v===null`＝★「マスが 空か」★を 見ていた
   ⇒ ISOMITTED は「★LAMBDA の 引数が 省かれたか★」を 見る 関数で、★別物★
     （実Excel … =ISOMITTED(空のマス) は ★False★／カンマで 省いた 時だけ True）
   ⇒★間違った 意味の 物を 名前だけ 残すと ★次に 触る 人が 使う★★ので 消した
   ⇒ 今の 作りは この ファイルの `ISOMITTED(` の 所（実測の 表を 隣に 書いた） */

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

/* ★範囲を「:」で つないだ物を 1つの四角に まとめる★（2026-08-29）
 *
 *  Excel の「:」は ★範囲演算子★＝2つの範囲を「囲む四角」に する。
 *  司さんの実物に こう書いてある（★実Excel(COM)で 原本の複製を 読んだ 真値★）:
 *      =SUM(ROUNDDOWN(FILTER(R8.1[見本一四]:R8.1[架空二]: … :R8.1[雛形八一], … )))
 *  表の名前を A1 に 直すと `'歩合'!V2:V32:'歩合'!X2:X32: … :'歩合'!AJ2:AJ32` になる。
 *
 *  ★HyperFormula は この形を 読めない★（実測 2026-08-29）:
 *      =SUM(A1:A3:C1:C3)              → Parsing error（RParen を待っていたら ':' が来た）
 *      =SUM('歩合'!A1:A3:'歩合'!C1:C3) → #ERROR!
 *      =SUM('歩合'!A1:C3)              → ★45（通る）★
 *  ⇒ ★囲む四角に まとめてから 渡す★。
 *
 *  これを入れる前は 司さんの実物で ★52本が #ERROR!★ だった
 *  （実Excel の答えは 177,523 / 211,558 / 279,558 / 668,639 ＝★金額★）。
 *
 *  ★同じシートの物だけ まとめる★（別シートを跨ぐ「:」は Excel でも 意味が違う）。
 *  ★読めない形は 触らない★（部品が 奇数個・番地の形が 違う 等）。
 */
function _mergeRangeChains(f) {
  var シート = "(?:'[^']+'|[A-Za-z_\\u3040-\\u30ff\\u4e00-\\u9fff][\\w\\u3040-\\u30ff\\u4e00-\\u9fff.]*)!";
  var 番地 = '\\$?[A-Za-z]{1,3}\\$?[0-9]{1,7}';
  var 一つ = '(?:' + シート + ')?' + 番地 + ':' + 番地;
  var 鎖 = new RegExp('(' + 一つ + ')(?::(' + 一つ + ')){1,}', 'g');
  return String(f).replace(鎖, function (whole) {
    var 部品 = whole.split(':');
    if (部品.length % 2 !== 0) return whole;        // 2つで1組。奇数なら 触らない
    var 名 = null, r1 = Infinity, r2 = -Infinity, c1 = Infinity, c2 = -Infinity;
    for (var i = 0; i < 部品.length; i++) {
      var m = 部品[i].match(/^(?:(.+)!)?\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/);
      if (!m) return whole;
      var このシート = m[1] || '';
      if (i === 0) 名 = このシート;
      else if (このシート && このシート !== 名) return whole;   // ★別シートは まとめない★
      var col = 0, L = m[2].toUpperCase();
      for (var k = 0; k < L.length; k++) col = col * 26 + (L.charCodeAt(k) - 64);
      var row = parseInt(m[3], 10);
      if (col < c1) c1 = col;
      if (col > c2) c2 = col;
      if (row < r1) r1 = row;
      if (row > r2) r2 = row;
    }
    function 列名(n) { var s = ''; while (n > 0) { var t = (n - 1) % 26; s = String.fromCharCode(65 + t) + s; n = (n - t - 1) / 26; } return s; }
    return (名 ? 名 + '!' : '') + 列名(c1) + r1 + ':' + 列名(c2) + r2;
  });
}

/* ★Excelが 新しい関数に 付ける 内部の印を 外す★（2026-08-29）
 *   Excel は 新しい関数を ファイルの中で `_xlfn.` `_xlws.` `_xludf.` `_xlpm.` 付きで 持つ。
 *   ★.xlsb を 読むと 裸の `_xlws.FILTER(` が そのまま 来る★（実測）。
 *   HyperFormula は これを 読めない ⇒ `Parsing error`。
 *   ★長い方から 順に★（先に短い方を外すと `_xlfn.` だけ 消えて `_xlws.` が 残る）。
 *   ★同じ一覧を lib/xlsx-io.js の stripXlfn も 持つ★（あちらは 式バーの表示用）。
 *     tests/xlfn-strip.test.mjs が ★両方が 同じ物を 外すか★ を 突き合わせる。 */
function _stripXlPrefix(f) {
  return String(f)
    .replace(/_xlfn\._xlws\./g, '')
    .replace(/_xlfn\./g, '')
    .replace(/_xlws\./g, '')
    .replace(/_xludf\./g, '')
    .replace(/_xlpm\./g, '');
}

/* ★★指数の 大文字 E を 小文字 e に する（2026-09-10）★★
 *
 *  ★何が 起きて いたか★
 *    お客さんが ★指数の 字を 打つと #ERROR★ に なって いました。
 *      =1.64E-14 ／ =1E+20 ／ =1.5E+20 ／ =2E5 ／ =1E-5 ／ =1E-4 … 全部 ★#ERROR★
 *      ★=1.64e-14（小文字）は 通る★
 *    ⇒★エンジンの 式読みが ★大文字 E を 数と 読めない★★
 *    ⇒★★プラグインでは 直りません（式を 読む 所＝借り物の 中）★★
 *      ⇒★借り物の 中は 読まず、★渡す 前に 書き換える★★
 *
 *  ★「E14 という セル参照と ぶつかって いる」は ★外れ★でした★
 *    押して 確かめました … `=E14` → 999 ／ `=SUM(E1:E14)` → 1002 ★どちらも 正しく 動く★
 *    ⇒ ぶつかって いません。★大文字 E を 数として 読まないだけ★です
 *
 *  ★★一番 危ないのは ★セル参照を 壊す 事★★★
 *    ⇒★壊れる 側から 試験を 書きました★（下の 見張り）
 *    ⇒ 触らない 物 … `=E14` `=SUM(E1:E14)` `=$E$14` `=Sheet1!E14`
 *                    `="1.64E-14"`（★字の 中★） `=TEXT(A1,"0.00E+00")`（★書式の 中★）
 *                    `=A1E5`（★前に 英字＝セルの 名前かも★） `=EXP(2)`（関数名）
 *
 *  ★見て いない 範囲★
 *    ・★出る 字は 直して いません★（うち `1.64e-14` ／ 実Excel `1.64E-14`）＝★別の 直し★
 *    ・★一重引用符（'）の 中は 見て いません★（シート名の 囲み）
 */
function _指数のEを小文字に(f){
  if(!f||f.indexOf('E')<0) return f;
  var 配列=f.split(''), 字の中=false;
  for(var i=0;i<配列.length;i++){
    if(配列[i]==='"'){ 字の中=!字の中; continue; }
    if(字の中) continue;                              /* ★字の 中は 触らない★ */
    if(配列[i]!=='E') continue;
    /* ★E の 前が 数字★（1.64E / 1E / 2E …） */
    var 前=配列[i-1];
    if(!(前>='0'&&前<='9')) continue;
    /* ★その 数字の 頭まで 戻り、その 前に 英字が 在れば セルの 名前★（A1E5 など） */
    var j=i-1;
    while(j>=0&&配列[j]>='0'&&配列[j]<='9') j--;
    if(j>=0&&配列[j]==='.'){ j--; while(j>=0&&配列[j]>='0'&&配列[j]<='9') j--; }
    if(j>=0&&/[A-Za-z_$]/.test(配列[j])) continue;
    /* ★E の 後ろ＝符号? ＋ 数字★（後ろに 英字が 続けば セルの 名前） */
    var k=i+1;
    if(配列[k]==='+'||配列[k]==='-') k++;
    var 数=0;
    while(k<配列.length&&配列[k]>='0'&&配列[k]<='9'){ k++; 数++; }
    if(!数) continue;
    if(k<配列.length&&/[A-Za-z_]/.test(配列[k])) continue;
    配列[i]='e';
  }
  return 配列.join('');
}

function convertFormula(f, シート) {
  if(!f||f[0]!=='=') return f;
  /* ★指数の 大文字 E は エンジンが 読めない★（上の 断り） */
  f = _指数のEを小文字に(f);
  /* ★名前の定義を 実際の範囲に 開く★（2026-08-29）
     HyperFormula は ★日本語の名前を 受け付けない★（実測：売上／うりあげ／売上_2026 とも invalid）。
     司さんの実物は 日本語の名前で 書かれているので ★エンジンに 入る前に 開く★。
     開き方は lib/named-ranges.js（字の中は 触らない／長い名前から 先に 開く）。 */
  if(typeof window !== 'undefined' && window.名前の箱 && typeof window.名前の箱.開く === 'function'){
    f = window.名前の箱.開く(f, シート);
  }
  if(f.indexOf('_xl') >= 0) f = _stripXlPrefix(f);
  if(f.indexOf(':') >= 0) f = _mergeRangeChains(f);
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
  //  ★YEN( → DOLLAR( … 同じ理由(第3波P2で実測)。通貨書式の関数の本名は DOLLAR で、
  //    YEN は日本語UIの表示名。Excel自身が .FormulaLocal(=YEN) ↔ .Formula(=DOLLAR) で相互変換する。
  //    日本語環境の DOLLAR は ¥ を出すので、YEN と打った人の期待どおりになる。
  f = f.replace(/\bYEN\s*\(/gi, 'DOLLAR(');
  f = f.replace(/\bFALSE\b(?!\s*\()/g, 'FALSE()');
  f = f.replace(/\bTRUE\b(?!\s*\()/g, 'TRUE()');
  f = f.replace(/\bNORMSDIST\s*\(([^,)]+)\)/gi, 'NORMSDIST($1,TRUE())');
  f = f.replace(/\bBETADIST\s*\(([^,()]+,[^,()]+,[^,()]+)\)/gi, 'BETA.DIST($1,TRUE())');
  f = f.replace(/\bHYPGEOMDIST\s*\(([^()]+)\)/gi, function(m,args){return args.split(',').length===4?'HYPGEOMDIST('+args+',FALSE())':m;});
  f = f.replace(/\bNEGBINOMDIST\s*\(([^()]+)\)/gi, function(m,args){return args.split(',').length===3?'NEGBINOMDIST('+args+',FALSE())':m;});
  f = _rewriteIsRef(f);
  f = _rewriteAreas(f);
  return f;
}

/* ★AREAS＝『範囲が 幾つに 分かれているか』を 数える★（2026-09-07）
 *
 *  ★★なぜ ここ（式の 字を 書き換える 段）で やるか★★
 *    エンジンは ★とびとびの 範囲を ★式を 読む 所★で 断ります★
 *      `=SUM((A1:B3,D1:D2))` … Parsing error（RParen を 待っていたら , が 来た）
 *    ⇒★関数の 所まで 届かない＝プラグインでは 受け取れない★
 *    ⇒★JS層も 駄目★＝★式の 一番 外側でしか 効かない★
 *      （実測 … AREAS は 入れ子で 使われる
 *        `=SUM(AREAS((A1,A2)),1)` → 3 ／ `=AREAS((A1,A2))+AREAS(B2:D4)` → 3）
 *    ⇒★★字を 数に 置き換えれば 入れ子も そのまま 正しく なる★★
 *      `=SUM(AREAS((A1,A2)),1)` → `=SUM(2,1)` → 3
 *
 *  ★答え★ 実Excel に 4回 打たせた ★89本★
 *    `docs/measured/kansuu46/golden-areas*-2026-09-07.tsv`
 *    ★決まり★＝★中身では なく ★マスを 指しているか★だけ★（ISREF と 同じ）
 *      `AREAS(IF(TRUE,D1,B1))` … D1 が `#DIV/0!` でも ★1★
 *
 *  ★★出さない 形★★（★半分 合う 答えを 出さない★）
 *    ★かっこが 入る 形は 全部★（`AREAS(INDEX(...))` ／ シート名の かっこ）
 *    ⇒ 実測 … `AREAS(INDEX(D1:D2,1))` は ★1★／`AREAS(INDEX(A1:C3,99,1))` は ★#REF!★
 *      ⇒★どちらも「計算したら 誤り」＝★値では 見分けられない★★
 *    ⇒★★今の 作りでは 出せません★＝`AREAS.MADA()` に して `#NAME?`（＝まだです）
 *      ★★そのまま 残しては いけない★★＝ISREF で 踏んだ
 *        （エンジンは AREAS を 知らないので #NAME? に なるが、
 *          ★とびとびは 式を 読む 所で 止まって #ERROR! に なる★）
 *
 *  ★`#NULL!` も 今は 出せません★
 *    実測 … エンジンの 誤りの 種類に ★NULL が 無い★
 *    ⇒ `AREAS(B2:D4 A1)`（重ならない 重なり）は ★まだ★に する
 *      （★#ERROR! で 出すと 実Excel と 字が 違う★）
 *
 *  見張り: tests/formula-areas.test.mjs
 */
function _rewriteAreas(f){
  if(!/\bAREAS\s*\(/i.test(f)) return f;
  var A = null;
  try{
    A = (typeof module === 'object' && module.exports && typeof require === 'function')
      ? require('./lib/formula-areas.js')
      : (typeof self !== 'undefined' ? self.FormulaAreas : null);
  }catch(e){ A = (typeof self !== 'undefined' ? self.FormulaAreas : null); }
  if(!A || typeof A.区画を数える !== 'function') return f;   /* 部品が 無い＝触らない */
  var 出='', i=0;
  while(i < f.length){
    var m = /\bAREAS\s*\(/i.exec(f.slice(i));
    if(!m){ 出 += f.slice(i); break; }
    var 頭 = i + m.index, 開 = 頭 + m[0].length;
    出 += f.slice(i, 頭);
    /* ★かっこを 数えて 閉じを 探す★（字の かたまりの 中の かっこは 数えない） */
    var 深=1, j=開, 字中=false;
    while(j < f.length){
      var c = f.charAt(j);
      if(字中){ if(c === '"') 字中 = false; }
      else if(c === '"') 字中 = true;
      else if(c === '(') 深++;
      else if(c === ')'){ 深--; if(!深) break; }
      j++;
    }
    if(深){ 出 += f.slice(頭); break; }               /* 閉じが 無い＝触らない */
    var 答 = A.区画を数える(f.slice(開, j));
    if(typeof 答 === 'number') 出 += String(答);
    /* ★★#VALUE! は エンジンが 持っている＝★そのまま 出す★★
       （実測 … `AREAS((Sheet1!A1,二枚目!A1))` は 実Excel も #VALUE!）
       ⇒★出せる 答えを「まだ」に しない★＝それも 嘘に なる */
    else if(答 && 答.誤り === 'VALUE') 出 += '#VALUE!';
    /* ★★#NULL! は エンジンに 無い★★（実測 … 誤りの 種類に NULL が 無い）
       ⇒ `#ERROR!` で 出すと ★実Excel と 字が 違う★
       ⇒★だから ここだけ「まだです」に する★ */
    else 出 += 'AREAS.MADA()';
    i = j + 1;
  }
  return 出;
}


/* ★ISREF＝『マスを 指しているか』を 見る 関数★（2026-09-07 に 作り直した）
 *
 *  ★★前の 作りは 壊れていました（本番に 在りました）★★
 *    `f.replace(/\bISREF\s*\(([^)]+)\)/gi, …)`
 *    ⇒★`[^)]+` は ★最初の ) で 止まります★＝★入れ子を 数えません★
 *      `=ISREF(SUM(A1:A2))` → `=FALSE())`      ←★閉じかっこが 1つ 多い＝#ERROR!★
 *      `=ISREF(IF(TRUE,A1,B1))` → `=FALSE(),A1,B1))`
 *    ⇒ さらに TRUE に する 形が ★A1 と A1:B2 だけ★でした
 *      `=ISREF(A:A)` `=ISREF(1:1)` `=ISREF(Sheet1!A1)` … ★実Excel は TRUE／うちは FALSE★
 *      ⇒★誤りも 出ない＝★黙って 逆の 答え★＝一番 見つけにくい★
 *
 *  ★★答えは 実Excel に 打たせて 取りました★★
 *    `docs/measured/kansuu46/golden-isref-2026-09-07.tsv`（25本）
 *    ★実Excel の 決まり★＝★中身では なく ★マスを 指しているか★だけ★
 *      `ISREF(E5)` … E5 が `#DIV/0!` でも ★TRUE★（★誤りは 伝わらない★）
 *      `ISREF(INDEX(A1:B2,1,1))` … ★TRUE★（INDEX は マスを 指す）
 *      `ISREF(SUM(A1:A2))` …… FALSE（SUM は 値）
 *      `ISREF(INDIRECT("zzz"))` … ★FALSE★（指せなかった）
 *      `ISREF(INDEX(A1:B2,99,1))` … ★FALSE★（外を 指した）
 *
 *  ★★今回は「嘘を 止める」所までです★★（指示役 2026-09-07・2a）
 *    ★字だけで ★確かに★ 決まる 形★ … TRUE / FALSE を 出す
 *    ★関数が 中に 入る 形★ ………… ★出しません★
 *      ⇒`INDIRECT("A1")` は TRUE ／`INDIRECT("zzz")` は FALSE
 *        ＝★指せたかは 字だけでは 決まりません★（AREAS と 同じ）
 *      ⇒★★間違った TRUE/FALSE より #NAME?（＝まだです）の 方が 良い★★
 *      ⇒ ISREF の まま 残す＝エンジンが #NAME? を 返す
 *    ★名前（なまえ）★ も 同じ＝★まだ 出しません★（名前の 一覧を 見ないと 決まらない）
 *
 *  見張り: tests/isref.test.mjs
 */
function _rewriteIsRef(f){
  if(!/\bISREF\s*\(/i.test(f)) return f;
  var 出='', i=0;
  while(i < f.length){
    var m = /\bISREF\s*\(/i.exec(f.slice(i));
    if(!m){ 出 += f.slice(i); break; }
    var 頭 = i + m.index;                 /* ISREF の 先頭 */
    var 開 = 頭 + m[0].length;            /* ( の 次 */
    出 += f.slice(i, 頭);
    /* ★かっこを ★数えて★ 閉じを 探す★（字の かたまりの 中の かっこは 数えない） */
    var 深 = 1, j = 開, 字中 = false;
    while(j < f.length){
      var c = f.charAt(j);
      if(字中){ if(c === '"') 字中 = false; }
      else if(c === '"') 字中 = true;
      else if(c === '(') 深++;
      else if(c === ')'){ 深--; if(!深) break; }
      j++;
    }
    if(深){ 出 += f.slice(頭); break; }   /* 閉じが 無い＝触らない（式を 壊さない） */
    var 中 = f.slice(開, j).trim();
    var 答 = _isRefKotae(中);
    /* ★★決められない 形は ★そのままに しては いけません★★（2026-09-07 実測）
       ⇒ エンジンは ★自分の ISREF を 持っています★
       ⇒ そのまま 渡すと ★エンジンが 答えを 出してしまう★
         実測 … `ISREF(INDIRECT("A1"))` → ★false★（実Excel は TRUE）
                `ISREF(INDIRECT("zzz"))` → ★true★（実Excel は FALSE）
                `ISREF(A1:INDEX(…))` → ★#ERROR!★
       ⇒★★『そのまま 残す＝黙る』では なく『別の 誰かが 嘘を 言う』でした★★
       ⇒ だから ★無い 名前★に 置き換えて ★#NAME?（＝まだです）★を 出す
         （お客さんに 見える 式は 元の まま＝書き換えた 字は 画面に 出ません） */
    出 += (答 === null) ? 'ISREF.MADA()'
                        : (答 ? 'TRUE()' : 'FALSE()');
    i = j + 1;
  }
  return 出;
}

/** ★字だけで 決まる 形か★ true／false／★null＝決められない★ */
function _isRefKotae(中){
  if(!中) return null;
  /* ★★TRUE()／FALSE() は かっこが 在っても ★値★★（2026-09-08 に 踏んだ）
     ⇒ `convertFormula` は ★この 関数より 先に★ `TRUE` を `TRUE()` に 直している
       （上の `f.replace(/\bTRUE\b(?!\s*\()/g, 'TRUE()')`）
     ⇒ だから ここへ 来る 時には ★もう かっこが 付いている★
     ⇒ 下の「かっこが 在れば 決められない」に 先に 引っかかり
       `=ISREF(TRUE)` が ★#NAME?★に なっていた（★実Excel は False★）
     ⇒★★#45 で 私が 本番に 入れた 戻り★★（★直す前は FALSE＝合っていた★）
     ⇒★『かっこが 在る＝関数』は 早すぎる 決めつけだった★
     ⇒★見つけ方★ 司さんの「答え確かめて」で 実Excel の 16本と 突き合わせて 出た */
  if(/^(TRUE|FALSE)\s*\(\s*\)$/i.test(中.trim())) return false;
  if(/[()]/.test(中)) return null;                 /* ★関数が 入る＝決められない★ */
  var s = 中.replace(/\$/g, '');
  /* ★シートの 名前が 付いていたら 外す★（'a b'!A1 も） */
  s = s.replace(/^'[^']*'!/, '').replace(/^[A-Za-z0-9_　-鿿]+!/, '');
  /* ★マスを 指す 形（実測で TRUE）★ */
  if(/^[A-Za-z]{1,3}[0-9]{1,7}(:[A-Za-z]{1,3}[0-9]{1,7})?$/.test(s)) return true;  /* A1 / A1:B2 */
  if(/^[A-Za-z]{1,3}:[A-Za-z]{1,3}$/.test(s)) return true;                          /* A:A */
  if(/^[0-9]{1,7}:[0-9]{1,7}$/.test(s)) return true;                                /* 1:1 */
  /* ★値そのもの（実測で FALSE）★ */
  if(/^"[^"]*"$/.test(中)) return false;                                            /* "a" */
  if(/^-?[0-9]+(\.[0-9]+)?$/.test(中)) return false;                                /* 1 / 1.5 */
  if(/^(TRUE|FALSE)(\(\))?$/i.test(中)) return false;                               /* TRUE */
  if(/[+\-*/&^]/.test(中)) return false;                                            /* 1+1 / A1&"" */
  return null;                                     /* ★名前など＝まだ 決めない★ */
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
  //  ★第3波P2(2026-08-01)で さらに11関数をここから外して HFプラグインへ移した:
  //    DOLLAR(旧YEN) / TYPE / AGGREGATE / LENB / LEFTB / RIGHTB / MIDB / RANK / VALUETOTEXT / ENCODEURL
  //  ★第3波P3(2026-08-02)で さらに11関数をここから外して HFプラグインへ移した:
  //    MODE / TRIMMEAN / PERCENTRANK / KURT / INTERCEPT / FORECAST / IRR /
  //    PERMUT / PERMUTATIONA / MDETERM / GESTEP
  //  ★残る5つ(PERCENTILE/QUARTILE/N/DSUM/DCOUNT)は「版上げで不要になる可能性」枠。判断日 2026-09-30。
  var _jsSet = {PERCENTILE:1,QUARTILE:1,
    XIRR:1,
    /* ★MIRR も 2026-09-09 に 足した★
       ★エンジン（借り物）が ★実Excel が 答える 所で #DIV/0! を 返して いた★
         =MIRR(全部マイナス,0.1,0.12) 実Excel ★−1★ ／ うち ★#DIV/0!★
         =MIRR(A1:A3,-1,0.12)        実Excel ★0.17132403714770583★ ／ うち ★#DIV/0!★
       ⇒★出来る 物を 止めて いた＝方針（全部 つける）に 反する★
       ⇒★借り物の 中は 読まず、JS層で 先に 受けて 自分で 出す★ */
    MIRR:1,
    /* ★IRR と XNPV も 2026-09-09 に 足した★
       ★エンジン（借り物）が 実Excel と 違う 答えを 返して いた★ので JS層で 先に 受ける
         =IRR(全部プラス) … 実Excel #NUM! ／ エンジン 16693338463349.688
         =XNPV(0,…)       … 実Excel #NUM! ／ エンジン 300
       ⇒★関所に 入れないと ★私の 段まで 届きません★（★実際 1回 素通りした★）★ */
    /* ★XNPV だけ 足した（2026-09-09）★
       ★IRR は 足しませんでした★＝`_PLUGIN_FUNCS` に 既に 在り（EX.IRR）、
       ★_jsSet にも 入れると 1つの 関数が 2か所に なる★
       ⇒ 見張り `tests/xlsx-harness/compare.mjs` が
         「★同じ 関数が 2箇所で 定義されている★」で 赤に して 教えて くれました
       ⇒★IRR は エンジン側（下の EX.IRR）で 直します★ */
    XNPV:1,
    DATESTRING:1,OFFSET:1,
    N:1,
    /* ★CONVERT は 2026-09-07 に ここから 外して lib/formula-yosoku-plug.js へ 移した★
       ⇒ ここに 居た 物は ★答えを 4桁で 丸めていた★（0.4536／実Excel 0.45359237）
       ⇒★1つの 関数は 1か所でだけ 定義する★（両方に 居ると 先に 当たった 方が 勝つ） */
    DSUM:1,DAVERAGE:1,DCOUNT:1,DCOUNTA:1,DMAX:1,DMIN:1,DPRODUCT:1,
    DGET:1,DSTDEV:1,DSTDEVP:1,DVAR:1,DVARP:1,
    /* ★LINEST は 2026-09-09 に 外しました★＝JS層で 受けなく なった
       （`lib/formula-yosoku.js` の `直線の係数()` が ★表として★ 返す）
       ⇒★関所に 残すと「JS層に 段が 在る」と 読まれる★ */
    BINOM:1,FREQUENCY:1,
    /* == 2026-09-19 ... SCAN / MAP / MAKEARRAY を ここから 外しました ==
         ★何が 起きて いたか★
           この 段（JS層）は ★エンジンより 先★に 答えます（記憶「Exally の 式は 2階建て」）。
           SCAN / MAP / MAKEARRAY は ここで ★先頭値だけ★ を 返して いました
           ⇒★★溢れる 式なのに 1マスしか 出ません★★
           ⇒経営者1 が 本番で 実測 ... 実Excel `6 / 2 / 4` ／ うち ★1x1（6 だけ）★
         ★なぜ 今まで 出なかったか★
           ★分母が 溢れの ★左上しか★ 見て いなかった★（`osu-dai-dake.mjs` 自身の 断り）
         ★BYROW / BYCOL は ここに 居ませんでした★
           ⇒★だから あの 2つだけ 溢れて いました★
           ⇒★同じ LAMBDA の 一族で 割れて いた 訳が これです★
         ★外して 良いと 決めた 訳（★測って から★）★
           ⑴台が ★同じ 答えを 出す★ ... 紙で ★8 / 8★
              MAP 2 ／ MAP(引数2つ) #VALUE! ／ SCAN 1 ／ MAKEARRAY(2,3) 1
              ／ MAKEARRAY(2,2) 1 ／ REDUCE 15 ／ BYROW 3 ／ BYCOL 11
           ⑵台の ★溢れた 先★が 実Excel と 同じ ... ★14 / 14★
              （`docs/measured/golden-kobore-zenmasu-2026-09-19.tsv`）
           ⇒★左上は 1本も 変わらず、★溢れた 先が 出る ように なります★★
         ★REDUCE と ISOMITTED は 残します★ ... ★溢れません★（1つの 値を 返す）
         ★下の 段（mScan / mMap / mMkarr）も 一緒に 消しました★
           ＝関所から 外した のに 段を 残すと `js-sou-sekisho` が
             ★「関所を 通らない 段が 在る」と 赤に します★（★門が 正しい★） */
    REDUCE:1,ISOMITTED:1};
  if(!_jsSet[_fnBase]) return null; // JS非対象 → HFへ

  var f = v.slice(1).trim().toUpperCase();
  var fOrig = v.slice(1).trim();

  // PERCENTILE / PERCENTILE.INC / PERCENTILE.EXC
  var mPerc=fOrig.match(/^PERCENTILE(?:\.INC|\.EXC)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([0-9.]+)\s*\)$/i);
  if(mPerc){var vals=_getRangeVals(sheet,mPerc[1]);if(vals.length)return String(_jsPercentile(vals,parseFloat(mPerc[2])));}

  // QUARTILE / QUARTILE.INC / QUARTILE.EXC
  var mQuart=fOrig.match(/^QUARTILE(?:\.INC|\.EXC)?\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([0-4])\s*\)$/i);
  if(mQuart){var vals=_getRangeVals(sheet,mQuart[1]);if(vals.length)return String(_jsQuartile(vals,parseInt(mQuart[2])));}

  /* ★死にコードを 12か所 消した（2026-09-08）★＝入口の 関所を 通らない 段だった。訳と 名簿＝docs/measured/shini-code-2026-09-08.md */







  // DATESTRING
  var mDs=fOrig.match(/^DATESTRING\s*\(([A-Z]+\d+|[0-9]+)\)$/i);
  if(mDs){var sv=_getSingleVal(sheet,mDs[1])||parseInt(mDs[1]);return _jsDatestring(sv);}

  // OFFSET
  var mOff=fOrig.match(/^OFFSET\s*\(([A-Z]+\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*\)$/i);
  if(mOff){var v=_jsOffset(sheet,mOff[1],parseInt(mOff[2]),parseInt(mOff[3]));return String(v);}


  // N
  var mN=fOrig.match(/^N\s*\(([^)]+)\)$/i);
  if(mN){var sv=_getSingleVal(sheet,mN[1]);return String(_jsN(sv!==null?sv:mN[1]));}

  // CONVERT … ★2026-09-07 に ここから 外した★（lib/formula-yosoku-plug.js へ）
  //   ここに 在った 物は 答えを 4桁で 丸めていた ⇒ =CONVERT(1,"lbm","kg") が ★0.4536★
  //   実Excel は ★0.45359237★／単位も 足りなかった（kibyte が #N/A）


  // REDUCE(initial, range, LAMBDA(...))
  var mReduce=fOrig.match(/^REDUCE\s*\(([^,]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*(LAMBDA\s*\(.+\))\s*\)$/is);
  if(mReduce){var reduceInit=parseFloat(mReduce[1].trim())||0;var reduceR=_jsReduceCompute(sheet,reduceInit,mReduce[2],mReduce[3]);if(reduceR!==null)return reduceR;}

  /* == 2026-09-19 ... SCAN / MAP / MAKEARRAY の 段を 消しました ==
       ★訳は 上の 関所（_jsSet）の 注に 書いて あります★
       ＝★溢れる 式を ここで 1マスに 潰して いました★
       ⇒★台（lib/shiki-hyou.js）が 溢れごと 出します★
       ★中の 関数（_jsScanCompute / _jsMapCompute / _jsMakearrayCompute）は 残して います★
         ＝`__test` の 口から 試験が 呼んで います（★口を 黙って 消さない★）
         ＝★お客さんの 道からは もう 呼ばれません★ */

  /* ★★ISOMITTED … ★『空の マス』と『省かれた 引数』は 別物★★（2026-09-08 実Excel 実測）
     ★前は こう だった★ `_getSingleVal` が null なら true
       ⇒★「マスが 空か」を 見ていた★／ISOMITTED は「★LAMBDA の 引数が 省かれたか★」を 見る 関数
       ⇒ =ISOMITTED(A1) も =ISOMITTED(2) も =ISOMITTED("あ") も ★true★ を 返していた
       ⇒★実Excel は 16本 とも False★（金の紙 golden-346-2026-09-08.tsv）
     ★実Excel に 聞いた（docs/measured/kansuu46/toru-isomitted.ps1）★
       =ISOMITTED(C1)（★空の マス★）               → ★False★
       =LAMBDA(x,ISOMITTED(x))(C1)（空の マス）    → ★False★
       =LAMBDA(x,y,ISOMITTED(y))(1,2)（渡した）   → False
       =LET(x,1,ISOMITTED(x))                    → False
       ★=LAMBDA(x,y,ISOMITTED(y))(1,)（カンマで 省く）→ True★  ←★ここだけ True★
       ★=LAMBDA(x,y,ISOMITTED(x))(,2)★                → True★
       ★=LAMBDA(x,y,z,ISOMITTED(y))(1,,3)★            → True★
       =LAMBDA(x,y,ISOMITTED(y))(1)（数が 足りない）  → ★#VALUE!★
       =ISOMITTED()（裸）                          → ★実Excel が 式ごと 受け付けない★
     ⇒★★引数が 書いてあれば 実Excel は ★必ず False★★（空の マスでも）
     ★カンマで 省いた 形は ★まだ★★＝`_jsLambdaExpand` が `ISOMITTED(())` に 展開し、
       下の 型（かっこの 中に かっこ）に 当たらないので ★ここを 素通りして 誤りに なる★
       ⇒★『まだ』は 誤りで 止まる＝★黙って 逆を 返さない★★（本番の 今も #ERROR）
       ⇒ 直すのは ★別の 押し★（`_jsLambdaExpand` を 触る＝LAMBDA を 取る 6個 全部に 効く）
     ★大文字★ … 画面には ★この 字が そのまま 出る★（book.html … cell.d = jsResult）
       ⇒ 実Excel の 画面は ★FALSE（大文字）★（tests/boolean.test.mjs 2026-08-30 実測）
       ⇒★前は 小文字 'false' を 返していた＝★出た 字が 違った★★
       ⇒ このファイルで 小文字を 返していたのは ★ここ 1か所だけ★（他は 全部 'TRUE'/'FALSE'） */
  var mIso=fOrig.match(/^ISOMITTED\s*\(([^)]*)\)$/i);
  if(mIso){
    if(!mIso[1].trim()) return null;   /* =ISOMITTED() … 実Excel が 受け付けない ⇒ 答えを 出さない */
    return 'FALSE';
  }

  // DSUM / DAVERAGE / DCOUNT / DCOUNTA / DMAX / DMIN / DPRODUCT / DGET / DSTDEV / DSTDEVP / DVAR / DVARP
  var mDb=fOrig.match(/^(DSUM|DAVERAGE|DCOUNT|DCOUNTA|DMAX|DMIN|DPRODUCT|DGET|DSTDEV|DSTDEVP|DVAR|DVARP)\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([^,]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*\)$/i);
  if(mDb){var funcMap={'DSUM':'SUM','DAVERAGE':'AVG','DCOUNT':'CNT','DCOUNTA':'CNTA','DMAX':'MAX','DMIN':'MIN','DPRODUCT':'PROD','DGET':'GET','DSTDEV':'STD','DSTDEVP':'STDP','DVAR':'VAR','DVARP':'VARP'};var fn=funcMap[mDb[1].toUpperCase()];var fieldArg=mDb[3].trim().replace(/^["']|["']$/g,'');var fieldVal=parseInt(fieldArg)||fieldArg;return String(_jsDbFunc(fn,sheet,mDb[2],fieldVal,mDb[4]));}

  // ★★LINEST は ここで 受けません（2026-09-09 に 外しました）★★
  //   ★前の 形★ 裸の =LINEST(範囲,範囲) だけを 拾い ★傾き 1つ★を 返して いた。
  //     ⇒ `=INDEX(LINEST(…),1,1)` は ★#NAME?★（エンジンに 積んで いなかった）
  //     ⇒★x が 2本以上（重回帰）で ★静かに 違う 答え★★
  //         =LINEST(G1:G6,H1:I6) … 実Excel 0.7708333333333329 ／ うち ★7.390243902439025★
  //     ⇒ LINEST は ★表を 返す 関数★なのに ★1つの 数★しか 返せない 場所だった
  //   ★今★ `lib/formula-yosoku.js` の `直線の係数()` が ★表として★ 返します。
  //     ・LOGEST と ★同じ 土台（回帰）★＝重回帰も 統計の 5行も 出せる
  //     ・実測 … docs/measured/golden-linest-hyou-2026-09-09.tsv（実Excel 119本）
  //   ⇒★ここで 受けると エンジンより 先に 答えて しまう★ので ★何も しません★
  //   （下請けの `_jsLinest` `_jsLogest` `_jsTrend` `_jsGrowth` は ★2026-09-11 に 消しました★）



  // BINOM.DIST.RANGE
  var mBdr=fOrig.match(/^BINOM\.DIST\.RANGE\s*\(([0-9]+)\s*,\s*([0-9.]+)\s*,\s*([0-9]+)(?:\s*,\s*([0-9]+))?\s*\)$/i);
  // ★4桁の 丸めを 外した（2026-09-08）★ 12本 測って ★3本が ちょうど 0★ に なって いた
  //   実Excel 2.1426377248779138e-7 → 0 ／ 相対の ずれ 最大 1.807e-1
  //   ⇒ docs/measured/golden-marume-A4-binom-2026-09-08.tsv
  if(mBdr)return String(_jsBinomDistRange(parseInt(mBdr[1]),parseFloat(mBdr[2]),parseInt(mBdr[3]),mBdr[4]?parseInt(mBdr[4]):undefined));

  // FREQUENCY - 最初の値だけ返す（スピル）
  var mFreq=fOrig.match(/^FREQUENCY\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\)$/i);
  if(mFreq){var data=_getRangeVals(sheet,mFreq[1]),bins=_getRangeVals(sheet,mFreq[2]);var r=_jsFrequency(data,bins);return String(r[0]);}


  // IRR / XIRR (既存)
  // ★4桁の 丸めを 外した（2026-09-08）★ 実Excel 0.5478870809078217 が 0.5479 に なって いた（相対 2.4e-5）
  //   ⇒ docs/measured/golden-marume-A-2026-09-08.tsv（★お金の 利回り＝丸めては いけない★）
  /* ══ ★MIRR★（2026-09-09）══════════════════════════════
     ★★これは「断る」では なく「止めて いたのを 動かす」直しです★★
     ★前は エンジンが 実Excel の 答える 所で 断って いました★
       =MIRR(全部マイナス,0.1,0.12) 実Excel ★−1★ ／ うち ★#DIV/0!★
       =MIRR(A1:A3,-1,0.12)        実Excel ★0.17132403714770583★ ／ うち ★#DIV/0!★
     ⇒★出来る 物を 止めて いた＝方針（Excel の 最上級・全部 つける）に 反する★

     ★実Excel の 実測★（docs/measured/golden-okane-4kansuu-2026-09-09.tsv・69本）
       普通 …………………… 0.17132403714770583
       ★全部 マイナス★ …… ★−1★（★断らない★）
       全部 プラス ………… #DIV/0!
       全部 0 ……………… #DIV/0!
       1件だけ …………… #DIV/0!
       空・字が 混ざる …… 0.10000000000000009（★落として 計算する★）
       ★借りる 利率★ ……… −1／−0.5／−0.01／0／0.01 ★どれでも 同じ 答え★
                            ⇒★答えに 効かない★（負の 流れが 1件で 0期に 在るから）
       ★回す 利率★ ……… −1 だけ ★#DIV/0!★／−0.5 は 0／0 は 0.14017542509913805

     ★★式は 実Excel の 定義から 自分で 書きました（借り物は 読んで いません）★★
       MIRR ＝ ( −(正の 流れを 回す利率で 最後まで 育てた 合計)
                 ÷ (負の 流れを 借りる利率で 今に 引き戻した 合計) )^(1/(n−1)) − 1
       ⇒★手で 検算して 実Excel と 6通り 一致★してから 入れました */
  var mMirr=fOrig.match(/^MIRR\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*(-?[0-9.]+)\s*,\s*(-?[0-9.]+)\s*\)$/i);
  if(mMirr){
    var m生=_getRangeAll(sheet,mMirr[1]);
    var m借=Number(mMirr[2]),m回=Number(mMirr[3]);
    if(!isFinite(m借)||!isFinite(m回))return '#VALUE!';
    var m数=[];
    for(var mi=0;mi<m生.length;mi++){
      var mv=m生[mi];
      if(mv===null||mv===undefined||mv==='')continue;              /* ★空は 落とす★（実測） */
      if(typeof mv==='boolean')continue;
      if(typeof mv==='string'&&!isFinite(Number(mv)))continue;      /* ★字も 落とす★（実測） */
      m数.push(Number(mv));
    }
    var mn=m数.length;
    if(mn<2)return '#DIV/0!';                                       /* ★1件だけ★（実測） */
    var m負=0,m正=0,m負が在る=false;
    for(var mj=0;mj<mn;mj++){
      if(m数[mj]<0){ m負が在る=true; m負+=m数[mj]/Math.pow(1+m借,mj); }
      if(m数[mj]>0){ m正+=m数[mj]*Math.pow(1+m回,mn-1-mj); }
    }
    /* ★負の 流れが 1つも 無い＝割れない＝#DIV/0!★（実測 … 全部プラス・全部0） */
    if(!m負が在る||m負===0)return '#DIV/0!';
    if(!isFinite(m負)||!isFinite(m正))return '#DIV/0!';
    /* ★★回す 利率が −1＝#DIV/0!★★（実測）
       (1+回す利率) が 0 に なり ★正の 流れが 全部 0 に 潰れます★
       ⇒ 実Excel は #DIV/0!（うちは 直す 前 -0.16333997346592444 を 返して いた）
       ⇒★1回目 私は これを 見落として いました★＝★11通り 押して 1本だけ 違った★ */
    if(1+m回===0)return '#DIV/0!';
    var m比=-m正/m負;
    if(!(m比>=0))return '#NUM!';
    return String(Math.pow(m比,1/(mn-1))-1);
  }

  /* ★★見当（第3引数）も 受ける（2026-09-09）★★
     ★前は 2つの 範囲だけ★の 形でした ⇒ `=XIRR(範囲,範囲,見当)` は ここを 素通り
     ⇒ 実Excel は 見当が ★−1 以下★で ★#NUM!★ を 返しますが、うちは 数を 返して いた
        =XIRR(E1:E3,F1:F3,-0.9) … 実Excel #NUM! ／ うち 1.0011933213132864
     ⇒ 実測 … 見当 0.1 / 0 / 10 / 1e6 は ★値を 返す★／−1 と −2 は ★#NUM!★ */
  /* ★★IRR は ここで 受けません（2026-09-09）★★
     ★1回 ここに 段を 作りましたが 外しました★
     ⇒ IRR は `_PLUGIN_FUNCS` に 在り ★エンジン側（EX.IRR）で 既に 自分たちが 計算して います★
     ⇒ ここにも 置くと ★1つの 関数が 2か所★に なる（見張りが 赤に した）
     ⇒★直すのは エンジン側 1か所★（下の `statFn('exIrr', 'EX.IRR', …)`） */

  /* ══ ★XNPV★（2026-09-09）══════════════════════════════
     ★エンジンは 実Excel が 断る 所で 数を 返して いました★
       =XNPV(0,…)      実Excel #NUM! ／ うち ★300★
       =XNPV(-0.01,…)  実Excel #NUM! ／ うち ★305.022090126012★
       =XNPV(-0.99,…)  実Excel #NUM! ／ うち ★7847.371130821853★
       空が 混ざる・字が 混ざる … 実Excel #NUM! ／ うち #VALUE!（★誤りの 種類が 違う★）
     ★★利率の 境目は 刻んで 詰めました★★
       −2／−1.01／−1／−0.99／−0.5／−0.1／−0.01／−0.001／★0★ … 全部 #NUM!
       ★0.0001★ から 値（299.9501405358503）
       ⇒★境目は 0（0 は 断る）★
       ⇒★★XIRR の 見当（0 は 通る）とは 逆＝1つの 関数で 決めた 線を 家族に 広げない★★ */
  var mXnpv=fOrig.match(/^XNPV\s*\((-?[0-9.]+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)\s*\)$/i);
  if(mXnpv){
    var x利=Number(mXnpv[1]);
    var x値=_getRangeAll(sheet,mXnpv[2]),x日=_getRangeAll(sheet,mXnpv[3]);
    if(!isFinite(x利))return '#VALUE!';
    /* ★範囲の 大きさが 違う＝#NUM!★ */
    if(x値.length!==x日.length)return '#NUM!';
    /* ★★空も 字も #NUM!★★（★XIRR は 空を 0 と 見て 通す＝ここが 違う★） */
    for(var xi=0;xi<x値.length;xi++){
      var xv=x値[xi],xd=x日[xi];
      if(xv===null||xv===undefined||xv==='')return '#NUM!';
      if(xd===null||xd===undefined||xd==='')return '#NUM!';
      if(typeof xv==='boolean'||typeof xd==='boolean')return '#NUM!';
      if(!isFinite(Number(xv))||!isFinite(Number(xd)))return '#NUM!';
    }
    if(x値.length<1)return '#NUM!';
    /* ★日付が 昇順で ない＝#NUM!★ */
    for(var xj=1;xj<x日.length;xj++)if(Number(x日[xj])<Number(x日[xj-1]))return '#NUM!';
    /* ★利率が 0 以下＝#NUM!★（★刻んで 詰めた 境目★） */
    if(!(x利>0))return '#NUM!';
    var x0=Number(x日[0]),x和=0;
    for(var xk=0;xk<x値.length;xk++){
      x和+=Number(x値[xk])/Math.pow(1+x利,(Number(x日[xk])-x0)/365);
    }
    return String(x和);
  }

  var mXirr=fOrig.match(/^XIRR\s*\(([A-Z]+\d+:[A-Z]+\d+)\s*,\s*([A-Z]+\d+:[A-Z]+\d+)(?:\s*,\s*(-?[0-9.]+))?\s*\)$/i);
  // ★★実Excel と 同じ 条件で 同じ 誤りを 返す（2026-09-09）★★
  //   ★前は 断らずに 数を 返して いた★
  //     =XIRR(全部プラス)  … 実Excel #NUM! ／ うち ★375686054239746.3★
  //     =XIRR(日付が逆順)  … 実Excel #NUM! ／ うち ★-0.5002981524324954★
  //   ⇒★誤りに ならず 数が 出る＝お客さんは 気づけない★（★お金の 利回り★）
  //   ⇒★これは「機能を 減らす」では ありません＝★実Excel と 同じに する★★
  //   ⇒ 条件は ★私が 決めず 実Excel に 打たせました★
  //      docs/measured/golden-xirr-sakaime-2026-09-09.tsv（61本）
  if(mXirr){
    var 生値=_getRangeAll(sheet,mXirr[1]),生日=_getRangeAll(sheet,mXirr[2]);
    var 見当=(mXirr[3]===undefined||mXirr[3]==='')?undefined:Number(mXirr[3]);
    var 判=_xirrの門(生値,生日,見当);
    if(判)return 判;                       /* ★実Excel と 同じ 誤り★ */
    /* ★★門が 作った 組を そのまま 使う★★
       ★前は ここで `filter(typeof v==='number')` を して いました★
       ⇒★空の マスが 落ちて 日付と ずれる★＝門が 通しても 答えが 変わって いた
         AA1=-1000／AA2=★空★／AA3=1100 … 実Excel 0.21063382029533387 ／ うち ★0.46563424984948343★
       ⇒★門と 計算で ★同じ 組★を 見る★（★口裏を 合わせる★） */
    var 組=_xirrの組(生値,生日);
    return String(_jsXirr(組.値,組.日,見当));
  }

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
    _hfGetDisplay: _hfGetDisplay, _hf答えの型: _hf答えの型, _applyTextFormat: _applyTextFormat,
    addSheetToEngine: addSheetToEngine,
    // 統計
    _jsRank: _jsRank, _jsPercentile: _jsPercentile, _jsQuartile: _jsQuartile,
    _jsMode: _jsMode, _jsTrimmean: _jsTrimmean, _jsPercentrank: _jsPercentrank,
    _jsFrequency: _jsFrequency, _jsKurt: _jsKurt, _jsPermut: _jsPermut,
    _jsPermutationa: _jsPermutationa, _jsProb: _jsProb, _jsBinomDistRange: _jsBinomDistRange,
    // 回帰
    _jsSlope: _jsSlope, _jsIntercept: _jsIntercept, _jsForecast: _jsForecast,
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
    _jsLookup: _jsLookup, _jsXmatch: _jsXmatch,
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
    _jsGestep: _jsGestep, _jsDbFunc: _jsDbFunc,
    _jsEncodeUrl: _jsEncodeUrl, _jsAggregate: _jsAggregate,
    /* _jsIsomitted は 2026-09-08 に 消した（「マスが 空か」を 見ていた＝別物） */
    // LET/LAMBDA系
    _parseFuncArgs: _parseFuncArgs,
    _jsLetExpand: _jsLetExpand, _jsLambdaExpand: _jsLambdaExpand,
    _jsReduceCompute: _jsReduceCompute, _jsScanCompute: _jsScanCompute,
    _jsMapCompute: _jsMapCompute, _jsMakearrayCompute: _jsMakearrayCompute,
    // ルーター
    convertFormula: convertFormula,
    /* ★見張りが 中を1つずつ 試せるように 出す★（入口だけ見て 緑にしない） */
    _stripXlPrefix: _stripXlPrefix, _mergeRangeChains: _mergeRangeChains,
    _hf答えの型: _hf答えの型,
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
  'CONCAT','LOOKUP','XMATCH','INDIRECT','DATEVALUE','NUMBERVALUE','FIXED','ASC','DBCS','TEXTBEFORE','TEXTAFTER',
  // ★第3波P2(2026-08-01)  ※YEN は DOLLAR の日本語UI表示名なので登録しない(convertFormula で寄せる)
  'DOLLAR','TYPE','AGGREGATE','LENB','LEFTB','RIGHTB','MIDB','RANK','VALUETOTEXT','ENCODEURL',
  // ★第3波P3(2026-08-02) 統計・財務・行列
  'MODE','TRIMMEAN','PERCENTRANK','KURT','INTERCEPT','FORECAST','IRR','PERMUT','PERMUTATIONA','MDETERM','GESTEP'];
// 名前に「.」が入る別名は翻訳表だけ足す(内部名は EX.RANKEQ / EX.RANKAVG)
var _PLUGIN_ALIAS = { 'EX.RANKEQ': 'RANK.EQ', 'EX.RANKAVG': 'RANK.AVG',
  'EX.MODESNGL': 'MODE.SNGL', 'EX.FORECASTLINEAR': 'FORECAST.LINEAR' };
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
    ErrorType: pick('ErrorType'),
    ArraySize: pick('ArraySize')   // ★配列を返す関数が「出力の大きさ」を申告するのに要る(R19)
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
  var ArraySize = P.ArraySize;   // ★無くても動く(その場合は素の配列関数だけ従来どおり)。R19の申告に使う
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
  //  縦1列の配列として返す。
  //  ★中身が空の時は配列を作らずスカラー 0 を返す。理由は2つ:
  //    (1) SimpleRangeValue.onlyValues([]) は TypeError で落ちる(実測。=SORT(空範囲) で踏んだ)
  //    (2) 実Excelの =SORT(空範囲) / =UNIQUE(空範囲) は 0 を返す(実測 2026-08-02)
  function col(arr){
    if(!arr.length) return 0;
    return SRV.onlyValues(arr.map(function(v){ return [v]; }));
  }
  /* ★★R19: 配列を返す関数は「出力が何行何列になるか」を HyperFormula に申告しないと、
   *   素の =SORT(A1:A10) が #VALUE!(Cell range not allowed) になる。
   *   包んだ形(=SUM(SORT(...)))だけは動くので、入れ子のケースしか無かったハーネスでは見えなかった。
   *   実測(2026-08-02): 申告を足すと通り、しかも隣のセルへ正しく展開(スピル)される。
   *   展開先に既にデータがあると HF は #SPILL! を返して既存データを守る(Excelと同じ)。
   *   ここは「入力の範囲の高さ＝出力の高さ」で申告する(SORT/UNIQUE/FILTERはいずれも縦1列を返すため)。
   *   ★大きさが読めない時は 1x1 と申告する＝スカラー扱いになるだけで、黙って違う数字は出ない。 */
  function rangeHeight(argAst){
    try {
      if(argAst && argAst.type === 'CELL_RANGE' && argAst.start && argAst.end){
        return Math.abs(argAst.end.row - argAst.start.row) + 1;
      }
    } catch(e){}
    return 1;
  }
  /* ★幅も 読む★（FILTER は 形を保つので 幅の申告が 要る）。読めない時は 1（＝スカラー扱い） */
  function rangeWidth(argAst){
    try {
      if(argAst && argAst.type === 'CELL_RANGE' && argAst.start && argAst.end){
        return Math.abs(argAst.end.col - argAst.start.col) + 1;
      }
    } catch(e){}
    return 1;
  }
  function sizeFromFirstArg(ast){
    if(!ArraySize) return null;
    var args = (ast && ast.args) || [];
    return new ArraySize(1, rangeHeight(args[0]));
  }
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

  //  ★配列を返す関数の「出力の大きさ」。SORT/UNIQUE/FILTER はいずれも
  //    第1引数の範囲と同じ高さの縦1列を返すので、1つの実装で足りる。
  ExallyPlugin.prototype.exArraySize = function(ast){ return sizeFromFirstArg(ast) || new ArraySize(1,1); };
  /* ★FILTER は 形を保つ★＝出す大きさも 入力の 高さ×幅 で 申告する。
     幅1 と 申告したままだと、2列以上を 返しても ★1列しか 展開されない★（2026-08-29）。 */
  ExallyPlugin.prototype.exFilterArraySize = function(ast){
    if(!ArraySize) return null;
    var a = (ast && ast.args) || [];
    return new ArraySize(rangeWidth(a[0]), rangeHeight(a[0]));
  };

  /* ★SORT / UNIQUE も 形を保つ★（2026-08-29 実Excel 16.0 で 測り直した）
   *   A1:B4 = [[300,1],[100,2],[200,3],[100,2]] の 真値:
   *     =COLUMNS(SORT(A1:B4))            → ★2★   （★行ごと★ 並べ替え・列は そのまま）
   *     =ROWS(SORT(A1:B4))               → ★4★
   *     =INDEX(SORT(A1:B4),1,1)          → ★100★
   *     =INDEX(SORT(A1:B4,2),1,1)        → ★300★（2列目で 並べ替え）
   *     =INDEX(SORT(A1:B4,2,-1),1,1)     → ★200★（降順）
   *     =INDEX(SORT(A1:B4,1,1,TRUE),1,1) → ★1★  （by_col＝★列ごと★ 並べ替え）
   *     =ROWS(UNIQUE(A1:B4))             → ★3★  （★行の重複★を 消す）
   *     =COLUMNS(UNIQUE(A1:B4,TRUE))     → ★2★
   *     =ROWS(UNIQUE(A1:B4,FALSE,TRUE))  → ★2★  （1度だけ出る行）
   *   ★直す前は FILTER と 同じ作り★＝平らにして 縦1列で 返していたので
   *   2列以上で 中身も 形も 壊れていた（実測 10項目中 6件が 実Excelと 違った）。
   *
   *   ★1列の時は 今までどおり★（空を落として 縦1列）。
   *   ＝実Excel は 空を 残して 後ろに 置く（`=ROWS(SORT(D1:D4))` → 4）が、
   *     ★ここは 突き合わせ台 tests/xlsx-harness の 見本値が 決めている所★なので
   *     ★測っただけで 今回は 変えていない★（変えるなら 見本値ごと 取り直す）。 */
  function 並び比べ(a, b, dir){
    var A = (a===null||a===undefined||a==='');
    var B = (b===null||b===undefined||b==='');
    if(A && B) return 0;
    if(A) return 1;            // ★空は いつも 後ろ★（実Excelで実測）
    if(B) return -1;
    if(typeof a==='number' && typeof b==='number') return (a-b)*dir;
    return String(a).localeCompare(String(b))*dir;
  }
  function 転置(m){
    var out=[], 幅=(m[0]?m[0].length:0);
    for(var c=0;c<幅;c++){ var 行=[]; for(var r=0;r<m.length;r++) 行.push(m[r][c]); out.push(行); }
    return out;
  }
  ExallyPlugin.prototype.exSort = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.SORT'), function(range, idx, order, byCol){
      var e = firstErr(flat(range)); if(e) return e;
      var M = rect(range);
      var 幅 = M[0] ? M[0].length : 0;
      if(幅 <= 1){
        var arr = flat(range).filter(function(v){ return v!==null && v!==undefined && v!==''; });
        var d1 = (toNum(order)===-1) ? -1 : 1;
        arr.sort(function(a,b){ return 並び比べ(a,b,d1); });
        return col(arr);
      }
      var 列で = (byCol===true || toNum(byCol)===1);
      var 表 = 列で ? 転置(M) : M;
      var k = Math.max(1, Math.round(toNum(idx)||1)) - 1;
      var dir = (toNum(order)===-1) ? -1 : 1;
      表 = 表.slice().sort(function(x, y){ return 並び比べ(x[k], y[k], dir); });
      return rectOut(列で ? 転置(表) : 表);
    });
  };
  ExallyPlugin.prototype.exUnique = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.UNIQUE'), function(range, byCol, once){
      var e = firstErr(flat(range)); if(e) return e;
      var M = rect(range);
      var 幅 = M[0] ? M[0].length : 0;
      if(幅 <= 1 && !(byCol===true || toNum(byCol)===1)){
        /* ★1列は 今までどおり★（見本値が 決めている） */
        var arr = flat(range).filter(function(v){ return v!==null && v!==undefined && v!==''; });
        var seen = [], out = [];
        for(var i=0;i<arr.length;i++){
          var key = (typeof arr[i]) + ' ' + String(arr[i]);
          if(seen.indexOf(key)<0){ seen.push(key); out.push(arr[i]); }
        }
        if(once===true || toNum(once)===1){
          var 数={};
          for(var j=0;j<arr.length;j++){ var k2=(typeof arr[j])+' '+String(arr[j]); 数[k2]=(数[k2]||0)+1; }
          out = out.filter(function(v){ return 数[(typeof v)+' '+String(v)]===1; });
        }
        return col(out);
      }
      var 列で = (byCol===true || toNum(byCol)===1);
      var 一度だけ = (once===true || toNum(once)===1);
      var 表 = 列で ? 転置(M) : M;
      var 鍵 = 表.map(function(行){ return 行.map(function(v){ return (typeof v)+' '+String(v); }).join(''); });
      var 数え = {};
      for(var m=0;m<鍵.length;m++) 数え[鍵[m]] = (数え[鍵[m]]||0) + 1;
      var 見た = {}, 残り = [];
      for(var n=0;n<表.length;n++){
        if(見た[鍵[n]]) continue;
        見た[鍵[n]] = true;
        if(一度だけ && 数え[鍵[n]] !== 1) continue;
        残り.push(表[n]);
      }
      if(!残り.length) return 0;
      return rectOut(列で ? 転置(残り) : 残り);
    });
  };
  ExallyPlugin.prototype.exText = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.TEXT'), function(v, fmt){
      if(isErr(v)) return v;
      if(isErr(fmt)) return fmt;
      var raw = flat(v)[0];
      /* ★★空の マスは 0★★（2026-09-11 実Excel に 打たせて 測った）
         ★空の マス★と ★空の 字★は ★別物★です
           =TEXT(空のマス,"0.00") → ★"0.00"★   （0 として 書式を 掛ける）
           =TEXT("","0.00")      → ★""★       （字の まま）
           =TEXT("abc","0.00")   → "abc"      （数に ならない 字は そのまま）
           =TEXT(" ","0.00")     → " "
           =TEXT(TRUE,"0.00")    → "TRUE"
           =TEXT("12","0.00")    → "12.00"    （数に なる 字は 数として）
         ★見つけ方★＝司さんの 実物（計算の 板）を 1マスずつ 突き合わせた
           B299 `=TEXT(A299,"aaa")`（A299 は 空）… 実Excel「土」／うち「」
           （0日目＝1900/1/0 は 土曜） */
      /* ★★空の マスは ★記号★で 来ます★★（2026-09-11 実測）
         `unwrap` が 記号を `''` に するので、★空の 字と 見分けが 付かなく なります★。
         ⇒★記号の うちに 見ます★ */
      var 生 = (v && typeof v==='object' && v.data) ? [].concat.apply([], v.data)[0]
             : (Array.isArray(v) ? [].concat.apply([], v)[0] : v);
      if(typeof 生 === 'symbol') raw = 0;              /* ★空の マス＝0★ */
      /* ★真偽は 字の まま★（実測 … =TEXT(TRUE,"0.00") → 「TRUE」／うちは 1.00 に して いた） */
      if(typeof raw === 'boolean') return raw ? 'TRUE' : 'FALSE';
      var n = toNum(raw);
      if(n===null) return raw==='' ? '' : String(raw);
      /* ★和暦・経過時間は 書式の 台へ★（出せない＝#VALUE!・実測 -1.5 の `[h]:mm`） */
      var 台r = _和暦か経過なら台へ(n, String(fmt));
      if(台r !== undefined) return 台r === null ? err(ErrorType.VALUE) : 台r;
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
  /* ★FILTER は「行を残す／列を残す」＝形を保つ★（2026-08-29 実Excel 16.0 で 測り直した）
   *
   *   =SUM(FILTER(A1:B4,{1;0;0;1}))     → ★505★（縦の条件＝★行★を選ぶ・2列とも 残る）
   *   =COLUMNS(FILTER(A1:B4,{1;0;0;1})) → ★2★
   *   =SUM(FILTER(A1:B4,{1,0}))         → ★1000★（横の条件＝★列★を選ぶ）
   *   =SUM(FILTER(A1:A4,{2;0;0;0}))     → ★100★（0以外の数は 真）
   *   =SUM(FILTER(A1:A4,{-1;0;0;0}))    → ★100★（負の数も 真）
   *   =FILTER(A1:A4,{0;0;0;0},"なし")    → ★なし★
   *
   *   ★直す前は 平らにして 1つずつ 拾い、縦1列で 返していた★＝
   *   2列以上だと ★中身も 形も 壊れる★（実測 =FILTER(A1:B4,{1;0;0;1}) が [[100],[2]]・COLUMNS=1）。
   *   これで 司さんの実物の 月別合計が ★177,523 などの金額 → 0★ に なっていた。 */
  function rect(v){
    if(v && typeof v==='object' && v.data) return v.data.map(function(row){ return row.map(unwrap); });
    return [[unwrap(v)]];
  }
  function 真か(c){
    if(c===true) return true;
    if(typeof c==='number') return c !== 0;              // ★0以外は 真（負も 真）★実Excelで実測
    if(typeof c==='string') return c.toUpperCase()==='TRUE';
    return false;
  }
  function rectOut(rows2d){
    if(!rows2d.length || !rows2d[0].length) return 0;    // col() と同じ理由（空は スカラー）
    return SRV.onlyValues(rows2d);
  }
  ExallyPlugin.prototype.exFilter = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.FILTER'), function(range, cond, ifEmpty){
      var A = rect(range), C = rect(cond);
      var e = firstErr(flat(range)) || firstErr(flat(cond)); if(e) return e;
      var 高 = A.length, 幅 = A[0] ? A[0].length : 0;
      var C高 = C.length, C幅 = C[0] ? C[0].length : 0;
      var out = [];
      /* ★縦の条件（1列）＝行を選ぶ★ ／ ★横の条件（1行）＝列を選ぶ★
         どちらとも 決められない形は ★触らずに #VALUE!★（黙って違う数字を出さない） */
      if(C幅 === 1 && C高 === 高){
        for(var r=0;r<高;r++) if(真か(C[r][0])) out.push(A[r].slice());
      } else if(C高 === 1 && C幅 === 幅){
        var 残す = [];
        for(var c2=0;c2<幅;c2++) if(真か(C[0][c2])) 残す.push(c2);
        for(var r2=0;r2<高;r2++){
          var 行 = [];
          for(var k2=0;k2<残す.length;k2++) 行.push(A[r2][残す[k2]]);
          out.push(行);
        }
        if(!残す.length) out = [];
      } else {
        return err(ErrorType.VALUE);
      }
      if(!out.length || !out[0] || !out[0].length){
        if(ifEmpty===undefined || ifEmpty===null) return err(ErrorType.NA);
        return col(flat(ifEmpty));
      }
      return rectOut(out);
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

  // ═══ 第3波P2(2026-08-01) ここから ═══════════════════════════════
  //  真値は tests/xlsx-harness/cases/91-wave3-p2.json ＋ 実Excel(16.0.20228)の golden。
  ExallyPlugin.prototype.exDollar = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.DOLLAR'), function(num, dec){
      if(isErr(num)) return num;
      var n = toNum(flat(num)[0]);
      if(n===null) return err(ErrorType.VALUE);
      return _jsDollar(n, (dec===undefined||dec===null) ? undefined : optNum(dec, 0));
    });
  };
  //  TYPE は「値の種類」を返すので、エラーも配列もそのまま受け取る必要がある。
  //  runFunction は既定でエラーを素通しさせてしまうため、AST を自分で評価する。
  ExallyPlugin.prototype.exType = function(ast, state){
    if(!ast.args || ast.args.length!==1) return err(ErrorType.NA);
    var v = this.evaluateAst(ast.args[0], state);
    if(v instanceof CellError) return 16;
    if(v && typeof v==='object' && typeof v.data !== 'undefined') return 64;   // SimpleRangeValue=配列
    if(Array.isArray(v)) return 64;
    if(typeof v==='number') return 1;
    if(typeof v==='boolean') return 4;
    if(typeof v==='symbol') return 1;                                          // 空セルは数値扱い
    if(typeof v==='string') return 2;
    return 1;
  };
  ExallyPlugin.prototype.exAggregate = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.AGGREGATE'), function(fnum, opts, ref, k){
      var f = toNum(flat(fnum)[0]), o = toNum(flat(opts)[0]);
      if(f===null || o===null) return err(ErrorType.VALUE);
      var A = flat(ref);
      //  オプション 2,3,6,7 は「エラーを無視」。それ以外はエラーがあればそのまま返す。
      var ignoreErr = (o===2||o===3||o===6||o===7);
      var e = firstErr(A);
      if(e && !ignoreErr) return e;
      if(ignoreErr) A = A.filter(function(v){ return !isErr(v); });
      var r = _jsAggregate(f, o, A, (k===undefined||k===null) ? undefined : toNum(flat(k)[0]));
      if(r==='#NUM!') return err(ErrorType.NUM);
      if(r==='#DIV/0!') return err(ErrorType.DIV_BY_ZERO);
      return fromJs(r);
    });
  };
  ExallyPlugin.prototype.exLenb = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.LENB'), function(s){
      if(isErr(s)) return s;
      return _jsLenb(txt(flat(s)[0]));
    });
  };
  function byteFn(name, fn, argc){
    ExallyPlugin.prototype[name] = function(ast, state){
      return this.runFunction(ast.args, state, this.metadata('EX.' + name.replace(/^ex/,'').toUpperCase()), function(s, a, b){
        if(isErr(s)) return s;
        var t = txt(flat(s)[0]);
        var n1 = optNum(a, argc===3 ? 1 : 1);
        var n2 = optNum(b, 1);
        if(n1 < 0 || n2 < 0) return err(ErrorType.VALUE);
        return fn(t, n1, n2);
      });
    };
  }
  byteFn('exLeftb',  function(t,n){ return _jsLeftb(t,n); }, 2);
  byteFn('exRightb', function(t,n){ return _jsRightb(t,n); }, 2);
  byteFn('exMidb',   function(t,st,ln){ return _jsMidb(t,st,ln); }, 3);
  function rankFn(name, key, avg){
    ExallyPlugin.prototype[name] = function(ast, state){
      return this.runFunction(ast.args, state, this.metadata(key), function(val, ref, order){
        if(isErr(val)) return val;
        var v = toNum(flat(val)[0]);
        if(v===null) return err(ErrorType.VALUE);
        var A = flat(ref);
        var e = firstErr(A); if(e) return e;
        return fromJs(_jsRank(v, A, (order===undefined||order===null) ? 0 : optNum(order,0), avg));
      });
    };
  }
  rankFn('exRank',    'EX.RANK',    false);
  rankFn('exRankEq',  'EX.RANKEQ',  false);
  rankFn('exRankAvg', 'EX.RANKAVG', true);
  ExallyPlugin.prototype.exValuetotext = function(ast, state){
    if(!ast.args || !ast.args.length) return err(ErrorType.NA);
    var v = this.evaluateAst(ast.args[0], state);
    var fmt = 0;
    if(ast.args.length > 1){
      var f = this.evaluateAst(ast.args[1], state);
      var n = toNum(unwrap(f));
      if(n!==null) fmt = n;
    }
    if(v instanceof CellError) return v;
    if(v && typeof v==='object' && typeof v.data !== 'undefined') v = flat(v)[0];
    return _jsValuetotext(unwrap(v), fmt);
  };
  ExallyPlugin.prototype.exEncodeurl = function(ast, state){
    return this.runFunction(ast.args, state, this.metadata('EX.ENCODEURL'), function(s){
      if(isErr(s)) return s;
      return fromJs(_jsEncodeUrl(txt(flat(s)[0])));
    });
  };
  // ═══ 第3波P2 ここまで ═══════════════════════════════════════════

  // ═══ 第3波P3(2026-08-02) ここから ═══════════════════════════════
  //  統計・財務・行列。真値は tests/xlsx-harness/cases/92-wave3-p3.json ＋ 実Excel(16.0.20228)。
  //  ★旧 _jsSet 側は答えを丸めていた(IRRとKURTを小数4桁で切っていた)。プラグインでは丸めない。
  function nums(v){ return flat(v).map(toNum).filter(function(x){ return x!==null && isFinite(x); }); }
  //  範囲を「行の配列」として取り出す(MDETERM は縦横の形が要る)
  function grid2d(v){
    if(v && typeof v==='object' && v.data) return v.data.map(function(row){ return row.map(function(x){ return toNum(unwrap(x)); }); });
    if(Array.isArray(v)) return v.map(function(row){ return (Array.isArray(row)?row:[row]).map(function(x){ return toNum(unwrap(x)); }); });
    return [[toNum(unwrap(v))]];
  }
  //  引数1つ(範囲)を取り、数値の並びを渡して1つの値を返す関数の共通形
  function statFn(name, key, calc){
    ExallyPlugin.prototype[name] = function(ast, state){
      return this.runFunction(ast.args, state, this.metadata(key), function(range, a2, a3){
        var A = flat(range);
        var e = firstErr(A); if(e) return e;
        var arr = nums(range);
        if(!arr.length) return err(ErrorType.NUM);
        return fromJs(calc(arr, a2, a3));
      });
    };
  }
  statFn('exMode', 'EX.MODE', function(arr){ return _jsMode(arr); });
  statFn('exModeSngl', 'EX.MODESNGL', function(arr){ return _jsMode(arr); });
  statFn('exKurt', 'EX.KURT', function(arr){ return _jsKurt(arr); });
  statFn('exTrimmean', 'EX.TRIMMEAN', function(arr, pct){
    var p = optNum(pct, 0);
    if(p < 0 || p >= 1) return '#NUM!';
    return _jsTrimmean(arr, p);
  });
  statFn('exPercentrank', 'EX.PERCENTRANK', function(arr, x, sig){
    var v = optNum(x, null);
    if(v===null) return '#N/A';
    return _jsPercentrank(arr, v, (sig===undefined||sig===null) ? undefined : optNum(sig, 3));
  });
  /* ★★IRR … 実Excel と 同じ 条件で 同じ 誤りを 返す（2026-09-09）★★
     ★前は 断らずに 数を 返して いました★
       =IRR(全部プラス)   実Excel #NUM! ／ うち ★16693338463349.688★
       =IRR(全部マイナス) 実Excel #NUM! ／ うち ★16693338463349.688★
       =IRR(全部0)        実Excel #NUM! ／ うち 0.1
       =IRR(1件だけ)      実Excel #NUM! ／ うち 0.1
       =IRR(A1:A3,10)     実Excel 0.18881944173101095 ／ うち ★-2168772783.613375★
       =IRR(A1:A3,100)    実Excel 0.18881944173155896 ／ うち ★-336432261411301570★
     ⇒ 条件は ★私が 決めず 実Excel に 打たせました★
        docs/measured/golden-okane-4kansuu-2026-09-09.tsv（69本）
     ★実Excel は 見当が 何でも 同じ 答え★（−0.99〜100 の 11通りで 全部 0.1888…）
     ★空・字は 落として 計算する★（実測 0.10000000000000009＝−1000 と 1100 の 2件）
     ★1件だけ は #NUM!★（★XIRR は #N/A＝同じ 家族でも 違う★） */
  statFn('exIrr', 'EX.IRR', function(arr, guess){
    var 数=[];
    for(var i=0;i<arr.length;i++){
      var v=arr[i];
      if(v===null||v===undefined||v==='')continue;            /* ★空は 落とす★ */
      if(typeof v==='boolean')continue;
      if(typeof v==='string'&&!isFinite(Number(v)))continue;   /* ★字も 落とす★ */
      数.push(Number(v));
    }
    if(数.length<2)return '#NUM!';                             /* ★1件だけ★ */
    var 正=false,負=false;
    for(var j=0;j<数.length;j++){ if(数[j]>0)正=true; if(数[j]<0)負=true; }
    if(!正||!負)return '#NUM!';                                /* ★符号が 片方だけ★ */
    /* ★★期間を 日付に 直して XIRR の 落ち着かせ方を 使う★★
       ⇒ `_jsXirr` は ★−1 より 下へ 行かせない／落ち着かなければ やり直す★を 持って いる
       ⇒★同じ 計算を 2つ 書かない★ */
    var 日=[];
    for(var k=0;k<数.length;k++)日.push(k*365);
    var g=(guess===undefined||guess===null)?0.1:optNum(guess,0.1);
    if(!isFinite(g))return '#VALUE!';
    return _jsXirr(数,日,g);
  });
  //  ys と xs の2範囲を取る物
  function twoRangeFn(name, key, calc){
    ExallyPlugin.prototype[name] = function(ast, state){
      return this.runFunction(ast.args, state, this.metadata(key), function(a, b, c){
        var e = firstErr(flat(a)) || firstErr(flat(b)) || (c!==undefined && c!==null ? firstErr(flat(c)) : null);
        if(e) return e;
        return fromJs(calc(a, b, c));
      });
    };
  }
  twoRangeFn('exIntercept', 'EX.INTERCEPT', function(ys, xs){
    var Y = nums(ys), X = nums(xs);
    if(Y.length !== X.length || !Y.length) return '#N/A';
    return _jsIntercept(Y, X);
  });
  twoRangeFn('exForecast', 'EX.FORECAST', function(x, ys, xs){
    var v = toNum(flat(x)[0]); if(v===null) return '#VALUE!';
    var Y = nums(ys), X = nums(xs);
    if(Y.length !== X.length || !Y.length) return '#N/A';
    return _jsForecast(v, Y, X);
  });
  twoRangeFn('exMdeterm', 'EX.MDETERM', function(m){
    var g = grid2d(m);
    if(!g.length || g.some(function(r){ return r.length !== g.length || r.some(function(x){ return x===null; }); })) return '#VALUE!';
    return _jsMdeterm(g);
  });
  //  数を2つ取る物
  function twoNumFn(name, key, calc){
    ExallyPlugin.prototype[name] = function(ast, state){
      return this.runFunction(ast.args, state, this.metadata(key), function(a, b){
        if(isErr(a)) return a;
        var n = toNum(flat(a)[0]);
        var k = (b===undefined||b===null) ? null : toNum(flat(b)[0]);
        if(n===null) return err(ErrorType.VALUE);
        return fromJs(calc(n, k));
      });
    };
  }
  twoNumFn('exPermut', 'EX.PERMUT', function(n, k){
    if(k===null || n<0 || k<0 || k>n) return '#NUM!';
    return _jsPermut(Math.trunc(n), Math.trunc(k));
  });
  twoNumFn('exPermutationa', 'EX.PERMUTATIONA', function(n, k){
    if(k===null || n<0 || k<0) return '#NUM!';
    return _jsPermutationa(Math.trunc(n), Math.trunc(k));
  });
  twoNumFn('exGestep', 'EX.GESTEP', function(n, step){
    return _jsGestep(n, step===null ? 0 : step);   // ★しきい値を省いたら0（実Excelで =GESTEP(5) が 1 なのを確認）
  });
  // ═══ 第3波P3 ここまで ═══════════════════════════════════════════

  var ANY = { argumentType: T.ANY };
  var OPT = { argumentType: T.ANY, optionalArg: true };
  ExallyPlugin.implementedFunctions = {
    //  ★arraySizeMethod = 出力の大きさの申告(R19)。これが無いと素の =SORT(A1:A10) が #VALUE! になる。
    'EX.SORT':       { method: 'exSort',       parameters: [ANY, OPT, OPT, OPT], arrayFunction: true, arraySizeMethod: 'exFilterArraySize' },
    'EX.UNIQUE':     { method: 'exUnique',     parameters: [ANY, OPT, OPT],      arrayFunction: true, arraySizeMethod: 'exFilterArraySize' },
    'EX.TEXT':       { method: 'exText',       parameters: [ANY, ANY] },
    'EX.TEXTJOIN':   { method: 'exTextjoin',   parameters: [ANY, ANY, ANY], repeatLastArgs: 1 },
    'EX.INT':        { method: 'exInt',        parameters: [ANY] },
    'EX.MOD':        { method: 'exMod',        parameters: [ANY, ANY] },
    'EX.VALUE':      { method: 'exValue',      parameters: [ANY] },
    'EX.XLOOKUP':    { method: 'exXlookup',    parameters: [ANY, ANY, ANY, OPT, OPT, OPT] },
    'EX.FILTER':     { method: 'exFilter',     parameters: [ANY, ANY, OPT],      arrayFunction: true, arraySizeMethod: 'exFilterArraySize' },
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
    'EX.TEXTAFTER':   { method: 'exTextafter',   parameters: [ANY, ANY, OPT, OPT, OPT, OPT] },
    // ★第3波P2(2026-08-01)
    'EX.DOLLAR':      { method: 'exDollar',      parameters: [ANY, OPT] },
    //  TYPE はエラー(16)や配列(64)そのものを受け取る必要があるので AST を自分で評価する
    //  ＝ parameters は使わず、引数が計算済みで渡ってこない指定にする。
    'EX.TYPE':        { method: 'exType',        parameters: [ANY], doesNotNeedArgumentsToBeComputed: true },
    'EX.AGGREGATE':   { method: 'exAggregate',   parameters: [ANY, ANY, ANY, OPT] },
    'EX.LENB':        { method: 'exLenb',        parameters: [ANY] },
    'EX.LEFTB':       { method: 'exLeftb',       parameters: [ANY, OPT] },
    'EX.RIGHTB':      { method: 'exRightb',      parameters: [ANY, OPT] },
    'EX.MIDB':        { method: 'exMidb',        parameters: [ANY, ANY, ANY] },
    'EX.RANK':        { method: 'exRank',        parameters: [ANY, ANY, OPT] },
    'EX.RANKEQ':      { method: 'exRankEq',      parameters: [ANY, ANY, OPT] },
    'EX.RANKAVG':     { method: 'exRankAvg',     parameters: [ANY, ANY, OPT] },
    'EX.VALUETOTEXT': { method: 'exValuetotext', parameters: [ANY, OPT], doesNotNeedArgumentsToBeComputed: true },
    'EX.ENCODEURL':   { method: 'exEncodeurl',   parameters: [ANY] },
    // ★第3波P3(2026-08-02) 統計・財務・行列
    'EX.MODE':         { method: 'exMode',         parameters: [ANY] },
    'EX.MODESNGL':     { method: 'exModeSngl',     parameters: [ANY] },
    'EX.TRIMMEAN':     { method: 'exTrimmean',     parameters: [ANY, ANY] },
    'EX.PERCENTRANK':  { method: 'exPercentrank',  parameters: [ANY, ANY, OPT] },
    'EX.KURT':         { method: 'exKurt',         parameters: [ANY] },
    'EX.INTERCEPT':    { method: 'exIntercept',    parameters: [ANY, ANY] },
    'EX.FORECAST':     { method: 'exForecast',     parameters: [ANY, ANY, ANY] },
    'EX.FORECASTLINEAR': { method: 'exForecast',   parameters: [ANY, ANY, ANY] },   // FORECAST.LINEAR は同じ物の新しい名前
    'EX.IRR':          { method: 'exIrr',          parameters: [ANY, OPT] },
    'EX.PERMUT':       { method: 'exPermut',       parameters: [ANY, ANY] },
    'EX.PERMUTATIONA': { method: 'exPermutationa', parameters: [ANY, ANY] },
    'EX.MDETERM':      { method: 'exMdeterm',      parameters: [ANY] },
    'EX.GESTEP':       { method: 'exGestep',       parameters: [ANY, OPT] }
  };
  var tr = {};
  _PLUGIN_FUNCS.forEach(function(n){ tr['EX.'+n] = n; });
  Object.keys(_PLUGIN_ALIAS).forEach(function(k){ tr[k] = _PLUGIN_ALIAS[k]; });
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
