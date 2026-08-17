/* table-refs.js — ★「表の名前での参照」を、実際のA1範囲に直す★
 *
 * ═══ なぜ要るのか（2026-08-18 実物で確かめた）═══════════════════════════
 *   司さんの実物（代行計算表2026.xlsb・式15,126本）のうち ★11,669本★ が
 *   Table[列名] の形（構造化参照）で書かれている。これが1本残らず #ERROR になっていた。
 *
 *   ★原因は計算エンジンではない。読み込みライブラリだった★
 *     実Excel(COM)の真値   =INDEX(R8.1[白石正人], MATCH(B4, R8.1[日付], 0))
 *     SheetJS が返す式     =INDEX(Table1[#Data],  MATCH(B4, Table1[#Data], 0))
 *   ＝ ★表の名前も 列名も 消えて「表のぜんぶ」に化けている★。
 *     INDEX と MATCH が同じ範囲を指す、意味の壊れた式になる。
 *     計算エンジンをいくら直しても、届く前に情報が消えているので絶対に直らない。
 *
 *   SheetJS 0.20.3 の該当箇所（lib/xlsx.full.min.js。★このファイルは凍結中なので触らない★）:
 *     case "PtgList": f.push("Table"+w[1].idx+"[#"+w[1].rt+"]")
 *   読み取った {coltype, colFirst, colLast} を ★書き出す時に捨てている★。
 *
 * ═══ 形式ごとに壊れ方が違う（実Excelで作った見本で実測）═══════════════
 *   .xlsx / .xlsm … ★壊れない★。SheetJS は 売上T[金額] を そのままの文字で返す
 *                    → ここでは ★文字を読んでA1範囲に直す★ だけでよい
 *   .xlsb         … ★壊れる★。列が消えるので、元のバイト列から拾い直す必要がある
 *                    → 記録の中の PtgList（表id・行の種類・列番号）を並び順に拾い、
 *                      SheetJS が出した Table<id>[#<種類>] の N番目と入れ替える
 *
 * ═══ 壊すより断る ═══════════════════════════════════════════════════
 *   ★1セルでも辻褄が合わなければ、そのセルは直さない（元のまま＝#ERROR のまま）★
 *   合わせるのは次の3つ。1つでも外れたら そのセルは触らない:
 *     ① 拾った PtgList の数 ＝ SheetJS の Table<N>[#…] の数
 *     ② N番目どうしで 表id が一致する
 *     ③ N番目どうしで 行の種類（Data/All/Headers/Totals/…）が一致する
 *   ★①②③は SheetJS が出した文字自身が答えを持っている★＝当て推量にならない。
 *
 * 依存: ZipSurgeon（zip の取り出しだけ）。SheetJS には依存しない。
 * 対応: ブラウザ / Node 両方。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.TableRefs = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ══ 小道具 ═══════════════════════════════════════════════════════ */

  function colName(i) {                       // 0 → 'A'
    var s = '';
    i++;
    while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = (i - 1 - m) / 26; }
    return s;
  }
  function u32(b, p) { return (b[p] | (b[p + 1] << 8) | (b[p + 2] << 16) | (b[p + 3] << 24)) >>> 0; }
  function u16(b, p) { return b[p] | (b[p + 1] << 8); }

  /* 可変長の数（記録の番号・長さ）。xlsb はこれで前置きされる */
  function readVar(b, p, maxBytes) {
    var v = 0, shift = 0, n = 0;
    for (;;) {
      if (p + n >= b.length) return null;
      var x = b[p + n];
      v |= (x & 0x7f) << shift;
      n++;
      if ((x & 0x80) === 0) break;
      shift += 7;
      if (n >= maxBytes) return null;
    }
    return { value: v >>> 0, bytes: n };
  }

  /** 記録を歩く。★終端ぴったりで終わらなければ null（＝この部品は読まない）★ */
  function walkRecords(b) {
    var recs = [], p = 0;
    while (p < b.length) {
      var id = readVar(b, p, 2); if (!id) return null; p += id.bytes;
      var len = readVar(b, p, 4); if (!len) return null; p += len.bytes;
      if (p + len.value > b.length) return null;
      recs.push({ id: id.value, start: p, len: len.value });
      p += len.value;
    }
    return p === b.length ? recs : null;
  }

  /* ══ 行の種類 ═══════════════════════════════════════════════════════
   *  PtgList の flags の 2〜6bit。SheetJS の並びと同じ物を、うちの言葉で持つ。
   *  ★実Excelで作った見本の18ケースで1つずつ確かめた（tests/table-refs.test.mjs）★ */
  var BAND_BY_CODE = {
    0: 'data',          // [列名] / [#Data]
    1: 'all',           // [#All]
    2: 'headers',       // [#Headers]
    4: 'data',          // [[#Data],[列名]]
    6: 'headersData',   // [[#Headers],[#Data]]
    8: 'totals',        // [#Totals]
    12: 'dataTotals',   // [[#Data],[#Totals]]
    16: 'thisRow',      // [@列名]
  };
  /* SheetJS が文字に出す時の名前（Table1[#Data] の "Data" の所）。突き合わせに使う */
  var SJ_NAME_BY_CODE = {
    0: 'Data', 1: 'All', 2: 'Headers', 4: '?Data2', 6: '?DataHeaders',
    8: 'Totals', 12: '?DataTotals', 16: '?Current',
  };

  /* ══ 表の定義 ═══════════════════════════════════════════════════════
   *  { id, sheet, rw1, rw2, cl1, cl2, header, totals, cols[], name }
   *  rw/cl は 0 から数える。header/totals は行数（0 か 1）。 */

  /** .xlsb の表の部品（table*.bin）を読む。
   *  ★correctness に使うのは 先頭32バイトの固定の場所だけ★（rfx・id・見出し行数・合計行数）。
   *  名前と列名は「人に見せる為だけ」＝読めなくても直す方には響かない。 */
  function readTableBin(bytes) {
    var recs = walkRecords(bytes);
    if (!recs) return null;
    for (var i = 0; i < recs.length; i++) {
      var r = recs[i];
      if (r.id !== 343) continue;                       // BrtBeginList
      if (r.len < 32) return null;
      var b = bytes, s = r.start;
      var def = {
        rw1: u32(b, s), rw2: u32(b, s + 4), cl1: u32(b, s + 8), cl2: u32(b, s + 12),
        id: u32(b, s + 20), header: u32(b, s + 24), totals: u32(b, s + 28),
        cols: [], name: '',
      };
      if (def.rw2 < def.rw1 || def.cl2 < def.cl1) return null;
      if (def.header > 1 || def.totals > 1) return null;   // 見た事がない形は扱わない
      /* 名前（見せる為だけ）。読めなくても続ける */
      try {
        var cch = u32(b, s + 64);
        if (cch > 0 && cch < 256 && s + 68 + cch * 2 <= s + r.len) {
          var out = '';
          for (var k = 0; k < cch; k++) out += String.fromCharCode(u16(b, s + 68 + k * 2));
          def.name = out;
        }
      } catch (e) { /* 名前は無くてよい */ }
      /* 列名（見せる為だけ）。BrtBeginListCol の 28バイト目から */
      for (var j = i + 1; j < recs.length; j++) {
        if (recs[j].id !== 347) continue;
        try {
          var cs = recs[j].start, cc = u32(b, cs + 28), nm = '';
          if (cc > 0 && cc < 256 && 32 + cc * 2 <= recs[j].len) {
            for (var m = 0; m < cc; m++) nm += String.fromCharCode(u16(b, cs + 32 + m * 2));
          }
          def.cols.push(nm);
        } catch (e2) { def.cols.push(''); }
      }
      return def;
    }
    return null;
  }

  /** .xlsx の表の部品（table*.xml）を読む */
  function readTableXml(text) {
    var m = /<table[^>]*\sref="([A-Z]+\d+:[A-Z]+\d+)"[^>]*>/.exec(text) || /<table[\s\S]*?>/.exec(text);
    if (!m) return null;
    var head = m[0];
    function attr(n) { var a = new RegExp('\\s' + n + '="([^"]*)"').exec(head); return a ? a[1] : null; }
    var ref = attr('ref'); if (!ref) return null;
    var rr = /^([A-Z]+)(\d+):([A-Z]+)(\d+)$/.exec(ref); if (!rr) return null;
    function colNum(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; }
    var idAttr = attr('id');
    var def = {
      rw1: parseInt(rr[2], 10) - 1, rw2: parseInt(rr[4], 10) - 1,
      cl1: colNum(rr[1]), cl2: colNum(rr[3]),
      id: idAttr === null ? -1 : parseInt(idAttr, 10),
      header: attr('headerRowCount') === null ? 1 : parseInt(attr('headerRowCount'), 10),
      totals: attr('totalsRowCount') === null ? 0 : parseInt(attr('totalsRowCount'), 10),
      cols: [], name: attr('displayName') || attr('name') || '',
    };
    if (def.rw2 < def.rw1 || def.cl2 < def.cl1) return null;
    if (def.header > 1 || def.totals > 1) return null;
    var re = /<tableColumn\b[^>]*\sname="([^"]*)"/g, c;
    while ((c = re.exec(text))) def.cols.push(unescapeXml(c[1]));
    return def;
  }
  function unescapeXml(s) {
    return String(s).replace(/&lt;/g, '<').replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, '&');
  }

  /* ══ 表の参照 → A1 範囲 ═════════════════════════════════════════════
   *  ★ここが唯一の「範囲を決める場所」★（.xlsb も .xlsx も同じ関数を通す）
   *  返り値は文字列。作れない時は null（＝そのセルは直さない）。 */
  function toRange(def, opt) {
    if (!def) return null;
    var band = opt.band, r1, r2;
    var dataFirst = def.rw1 + def.header;
    var dataLast = def.rw2 - def.totals;
    if (dataLast < dataFirst) return null;                     // 中身が0行の表は扱わない

    if (band === 'data') { r1 = dataFirst; r2 = dataLast; }
    else if (band === 'all') { r1 = def.rw1; r2 = def.rw2; }
    else if (band === 'headers') { if (!def.header) return null; r1 = def.rw1; r2 = def.rw1 + def.header - 1; }
    else if (band === 'totals') { if (!def.totals) return null; r1 = def.rw2 - def.totals + 1; r2 = def.rw2; }
    else if (band === 'headersData') { if (!def.header) return null; r1 = def.rw1; r2 = dataLast; }
    else if (band === 'dataTotals') { if (!def.totals) return null; r1 = dataFirst; r2 = def.rw2; }
    else if (band === 'thisRow') {
      if (opt.curRow === undefined || opt.curRow === null) return null;
      /* ★自分の行が 中身の行の外なら 直さない★（Excel は #VALUE! を出す所） */
      if (opt.curRow < dataFirst || opt.curRow > dataLast) return null;
      r1 = r2 = opt.curRow;
    } else return null;

    var nCols = def.cl2 - def.cl1 + 1, c1, c2;
    if (opt.colFirst === undefined || opt.colFirst === null) { c1 = def.cl1; c2 = def.cl2; }
    else {
      if (opt.colFirst < 0 || opt.colFirst >= nCols) return null;
      var last = (opt.colLast === undefined || opt.colLast === null) ? opt.colFirst : opt.colLast;
      if (last < opt.colFirst || last >= nCols) return null;
      c1 = def.cl1 + opt.colFirst; c2 = def.cl1 + last;
    }

    var a = colName(c1) + (r1 + 1), b = colName(c2) + (r2 + 1);
    var range = (a === b) ? a : (a + ':' + b);
    /* ★別のシートに在る表なら シート名を付ける★（付けないと自分のシートを見て 黙って違う答えを出す）
       引用符は付けない＝book.html の quoteSheetRefs が必要な物だけ付ける（付け方を2か所に持たない） */
    if (opt.fromSheet && def.sheet && opt.fromSheet !== def.sheet) return def.sheet + '!' + range;
    return range;
  }

  /* ══ .xlsb: 記録から PtgList を拾う ══════════════════════════════════ */

  /** 式セルの記録の中で、式のトークン列(rgce)が何バイト目から何バイトかを返す。
   *  ★終端ぴったりに合わなければ null（＝この記録は読まない）★ */
  function rgceRange(id, b, s, len) {
    var p = 8;                                    // Cell = 列(4) + 書式/フラグ(4)
    if (id === 9) p += 8;                         // BrtFmlaNum: 数値(8)
    else if (id === 8) {                          // BrtFmlaString: 文字(4 + 2×文字数)
      if (s + p + 4 > s + len) return null;
      var cch = u32(b, s + p); if (cch > 0x7fffffff) return null;
      p += 4 + cch * 2;
    } else if (id === 10 || id === 11) p += 1;    // 真偽 / エラー
    else return null;
    p += 2;                                       // grbitFlags
    if (p + 4 > len) return null;
    var cce = u32(b, s + p); p += 4;
    if (p + cce + 4 > len) return null;
    var start = s + p;
    p += cce;
    var cb = u32(b, s + p); p += 4;
    if (p + cb !== len) return null;              // ★ぴったりでなければ断る
    return { start: start, len: cce };
  }

  /** rgce の中の PtgList を並び順に拾う。
   *  ★署名（0x18|0x38|0x58, 0x19）で探す。★正しさは呼ぶ側が SheetJS の文字と突き合わせて保証する★
   *  （式のトークンを全種類 数え直す作りにすると、1種類でも寸法を間違えた時に黙って化ける） */
  function findPtgLists(b, rg) {
    var out = [], end = rg.start + rg.len;
    for (var i = rg.start; i + 14 <= end; i++) {
      var b0 = b[i];
      if ((b0 !== 0x18 && b0 !== 0x38 && b0 !== 0x58) || b[i + 1] !== 0x19) continue;
      var flags = u16(b, i + 4);
      out.push({
        coltype: flags & 3,
        code: (flags >> 2) & 31,
        id: u32(b, i + 6),
        colFirst: u16(b, i + 10),
        colLast: u16(b, i + 12),
      });
      i += 13;
    }
    return out;
  }

  /* SheetJS が出す placeholder。Table<数>[#<名前>] */
  var PLACEHOLDER = /Table(\d+)\[#([A-Za-z?0-9]+)\]/g;

  /** SheetJS の式 + 拾った PtgList → 直した式（合わなければ null） */
  function rewriteXlsbFormula(formula, lists, tablesById, ctx) {
    if (typeof formula !== 'string' || formula.indexOf('[#') < 0) return null;
    PLACEHOLDER.lastIndex = 0;
    var hits = [], m;
    while ((m = PLACEHOLDER.exec(formula))) hits.push({ at: m.index, text: m[0], id: +m[1], rt: m[2] });
    if (!hits.length) return null;
    if (hits.length !== lists.length) return null;                 // ① 数が合わない → 直さない

    var out = '', prev = 0;
    for (var i = 0; i < hits.length; i++) {
      var h = hits[i], L = lists[i];
      if (h.id !== L.id) return null;                              // ② 表id が合わない
      if (SJ_NAME_BY_CODE[L.code] !== h.rt) return null;           // ③ 行の種類が合わない
      /* ★消された表への参照（id = 0xFFFFFFFF）は Excel 自身が #REF! と書いている★
         実物で確認（2026-08-18・実Excelの .Formula）:
           計算!AJ269 =IFERROR(([@売上合計]-…-#REF!)+…,"")      → 画面は空（IFERROR が拾う）
           給料3!B75  =INDEX(#REF!, MATCH(B74, R8.3[日付], 0))   → 画面は #REF!
         ここで断ると 式そのものが読めなくなり、★IFERROR まで道連れで #ERROR になる★。
         Excel と同じく #REF! を置く＝IFERROR が拾えるようにする。 */
      if (L.id === 0xFFFFFFFF) {
        out += formula.slice(prev, h.at) + '#REF!';
        prev = h.at + h.text.length;
        continue;
      }
      var def = tablesById[L.id];
      if (!def) return null;
      var opt = { band: BAND_BY_CODE[L.code], curRow: ctx.row, fromSheet: ctx.sheet };
      if (L.coltype === 1) { opt.colFirst = L.colFirst; opt.colLast = L.colFirst; }
      else if (L.coltype === 2) { opt.colFirst = L.colFirst; opt.colLast = L.colLast; }
      else if (L.coltype !== 0) return null;
      var range = toRange(def, opt);
      if (!range) return null;
      out += formula.slice(prev, h.at) + range;
      prev = h.at + h.text.length;
    }
    return out + formula.slice(prev);
  }

  /* ══ .xlsx / .xlsm: 文字の形の表参照を直す ════════════════════════════
   *  SheetJS は 売上T[金額] を そのままの文字で返す（実測で壊れていない）。
   *  ここでは その文字を読んで A1 範囲に直す。 */

  /* 列名の中の逃がし記号。Excel は ' で逃がす（'[ '] '# '' ） */
  function unescapeColName(s) { return String(s).replace(/'([\[\]#'])/g, '$1'); }

  /** 表参照ひとつ分を読む。text の pos から始まる '[' の対応を数えて切り出す */
  function readBracket(text, pos) {
    if (text.charAt(pos) !== '[') return null;
    var depth = 0, i = pos, inQuote = false;
    for (; i < text.length; i++) {
      var ch = text.charAt(i);
      if (inQuote) { inQuote = false; continue; }     // ' の次の1文字は素通し
      if (ch === "'") { inQuote = true; continue; }
      if (ch === '[') depth++;
      else if (ch === ']') { depth--; if (!depth) return { end: i + 1, body: text.slice(pos + 1, i) }; }
    }
    return null;
  }
  /** [a],[b] を , で分ける（' の逃がしと [ ] の入れ子を数える） */
  function splitParts(body) {
    var out = [], cur = '', depth = 0, inQuote = false;
    for (var i = 0; i < body.length; i++) {
      var ch = body.charAt(i);
      if (inQuote) { cur += ch; inQuote = false; continue; }
      if (ch === "'") { cur += ch; inQuote = true; continue; }
      if (ch === '[') { depth++; cur += ch; continue; }
      if (ch === ']') { depth--; cur += ch; continue; }
      if (ch === ',' && !depth) { out.push(cur); cur = ''; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }
  var BAND_BY_TEXT = {
    '#all': 'all', '#headers': 'headers', '#data': 'data', '#totals': 'totals', '#this row': 'thisRow',
  };

  /** 表参照の中身（[] の内側）を読んで {band, colFirst, colLast} にする。読めなければ null */
  function parseRefBody(body, def, atMark) {
    var band = atMark ? 'thisRow' : null, cols = [];
    var parts = splitParts(body);
    /* [[金額]:[時間]] のような 列の範囲は、部品が1つで中に ':' が在る形になる */
    if (parts.length === 1) {
      var one = parts[0];
      var colon = topLevelColon(one);
      if (colon >= 0) {
        var a = stripBrackets(one.slice(0, colon)), b = stripBrackets(one.slice(colon + 1));
        if (a === null || b === null) return null;
        cols = [a, b];
        parts = [];
      }
    }
    for (var i = 0; i < parts.length; i++) {
      var p = parts[i].trim();
      if (!p) continue;
      var inner = stripBrackets(p);
      var raw = inner === null ? p : inner;
      var low = raw.toLowerCase();
      if (raw.charAt(0) === '#' && BAND_BY_TEXT[low]) {
        var bnd = BAND_BY_TEXT[low];
        if (band === 'headers' && bnd === 'data') band = 'headersData';
        else if (band === 'data' && bnd === 'headers') band = 'headersData';
        else if (band === 'data' && bnd === 'totals') band = 'dataTotals';
        else if (band === 'totals' && bnd === 'data') band = 'dataTotals';
        else if (band === null || band === bnd) band = bnd;
        else return null;
      } else if (raw.charAt(0) === '@') {
        band = 'thisRow';
        var nm = raw.slice(1);
        if (nm) { var s2 = stripBrackets(nm); cols.push(unescapeColName(s2 === null ? nm : s2)); }
      } else {
        cols.push(unescapeColName(raw));
      }
    }
    if (band === null) band = 'data';
    var out = { band: band };
    if (cols.length) {
      var i1 = def.cols.indexOf(cols[0]);
      if (i1 < 0) return null;
      out.colFirst = i1;
      out.colLast = cols.length > 1 ? def.cols.indexOf(cols[1]) : i1;
      if (out.colLast < 0) return null;
    }
    return out;
  }
  function stripBrackets(s) {
    var t = String(s).trim();
    if (t.charAt(0) !== '[' || t.charAt(t.length - 1) !== ']') return null;
    return t.slice(1, t.length - 1);
  }
  function topLevelColon(s) {
    var depth = 0, inQuote = false;
    for (var i = 0; i < s.length; i++) {
      var ch = s.charAt(i);
      if (inQuote) { inQuote = false; continue; }
      if (ch === "'") { inQuote = true; continue; }
      if (ch === '[') depth++;
      else if (ch === ']') depth--;
      else if (ch === ':' && !depth) return i;
    }
    return -1;
  }

  /** 名前で表を引く（表の名前は大文字小文字を区別しない＝Excel と同じ） */
  function findByName(tables, name) {
    var low = String(name).toLowerCase();
    for (var i = 0; i < tables.length; i++) if (String(tables[i].name).toLowerCase() === low) return tables[i];
    return null;
  }

  /** .xlsx の式（文字が正しい）→ A1 範囲に直した式。直す所が無ければ null */
  function rewriteTextFormula(formula, tables, ctx, ownTable) {
    if (typeof formula !== 'string' || formula.indexOf('[') < 0) return null;
    var out = '', i = 0, changed = false;
    while (i < formula.length) {
      var ch = formula.charAt(i);
      if (ch === '"') {                                   // 文字列の中は触らない
        out += ch; i++;
        while (i < formula.length) { out += formula.charAt(i); if (formula.charAt(i) === '"') { i++; break; } i++; }
        continue;
      }
      if (ch === '[') {                                   // [@列名] ＝ 表の名前を省いた形
        var br0 = readBracket(formula, i);
        if (!br0 || !ownTable) return null;               // 表の外で [ が出たら 直さない
        var spec0 = parseRefBody(br0.body, ownTable, br0.body.charAt(0) === '@');
        if (!spec0) return null;
        spec0.curRow = ctx.row; spec0.fromSheet = ctx.sheet;
        var rg0 = toRange(ownTable, spec0);
        if (!rg0) return null;
        out += rg0; i = br0.end; changed = true;
        continue;
      }
      /* 名前 + [ …… ] の形 */
      if (/[A-Za-z_\\-￿]/.test(ch)) {
        var j = i;
        while (j < formula.length && /[A-Za-z0-9_.\\-￿]/.test(formula.charAt(j))) j++;
        var word = formula.slice(i, j);
        if (formula.charAt(j) === '[') {
          var def = findByName(tables, word);
          if (def) {
            var br = readBracket(formula, j);
            if (!br) return null;
            var spec = parseRefBody(br.body, def, false);
            if (!spec) return null;
            spec.curRow = ctx.row; spec.fromSheet = ctx.sheet;
            var rg = toRange(def, spec);
            if (!rg) return null;
            out += rg; i = br.end; changed = true;
            continue;
          }
        }
        out += word; i = j;
        continue;
      }
      out += ch; i++;
    }
    return changed ? out : null;
  }

  /* ══ 入口 ═══════════════════════════════════════════════════════════ */

  /** zip から 表の定義（どのシートの表か まで）を集める */
  function loadTables(zip, isXlsb) {
    var relRe = /Id="([^"]+)"[^>]*Target="([^"]+)"/g;
    var names = zip.names();
    var sheetParts = names.filter(function (n) {
      return isXlsb ? /^xl\/worksheets\/sheet\d+\.bin$/.test(n) : /^xl\/worksheets\/sheet\d+\.xml$/.test(n);
    });
    /* シート部品 → シート名（workbook の並びと rels から引く。★部品の番号順に頼らない★） */
    var wbPart = isXlsb ? 'xl/workbook.bin' : 'xl/workbook.xml';
    var wbRels = isXlsb ? 'xl/_rels/workbook.bin.rels' : 'xl/_rels/workbook.xml.rels';
    return Promise.all([
      zip.has(wbPart) ? zip.bytes(wbPart) : Promise.resolve(null),
      zip.has(wbRels) ? zip.text(wbRels) : Promise.resolve(''),
    ]).then(function (r) {
      var wbBytes = r[0], relsText = r[1] || '';
      var relTarget = {}, m;
      relRe.lastIndex = 0;
      while ((m = relRe.exec(relsText))) relTarget[m[1]] = m[2].replace(/^\/?xl\//, '').replace(/^\.\.\//, '');
      var sheetNameOfPart = {};
      var order = [];
      if (isXlsb && wbBytes) {
        var recs = walkRecords(wbBytes);
        if (recs) {
          for (var i = 0; i < recs.length; i++) {
            if (recs[i].id !== 156) continue;                   // BrtBundleSh
            var s = recs[i].start, b = wbBytes;
            var p = 8, c1 = u32(b, s + p), rel = null;
            if (c1 === 0xFFFFFFFF) p += 4;
            else { rel = ''; for (var k = 0; k < c1; k++) rel += String.fromCharCode(u16(b, s + p + 4 + k * 2)); p += 4 + c1 * 2; }
            var c2 = u32(b, s + p), nm = '';
            for (var k2 = 0; k2 < c2; k2++) nm += String.fromCharCode(u16(b, s + p + 4 + k2 * 2));
            order.push({ name: nm, rel: rel });
          }
        }
      } else if (wbBytes) {
        var txt = '';
        for (var t = 0; t < wbBytes.length; t++) txt += String.fromCharCode(wbBytes[t]);
        try { txt = decodeURIComponent(escape(txt)); } catch (e) { /* そのまま */ }
        var shRe = /<sheet\b[^>]*\sname="([^"]*)"[^>]*r:id="([^"]*)"/g, sm;
        while ((sm = shRe.exec(txt))) order.push({ name: unescapeXml(sm[1]), rel: sm[2] });
      }
      order.forEach(function (o) {
        var tgt = o.rel && relTarget[o.rel];
        if (tgt) sheetNameOfPart['xl/' + tgt.replace(/^xl\//, '')] = o.name;
      });
      /* ★workbook から引けなかったら 番号順で当てない＝表のシートが分からない事にする★
         （黙って別のシートの範囲を作るより、直さない方が安全） */

      /* シートごとの rels → 表の部品 */
      var jobs = sheetParts.map(function (sp) {
        var relName = sp.replace(/([^/]+)$/, '_rels/$1.rels');
        if (!zip.has(relName)) return Promise.resolve({ sheet: sheetNameOfPart[sp] || null, tables: [] });
        return zip.text(relName).then(function (tx) {
          var out = [], mm, re = /Target="([^"]*tables\/table\d+\.(?:bin|xml))"/g;
          while ((mm = re.exec(tx))) out.push('xl/' + mm[1].replace(/^\.\.\//, '').replace(/^\/?xl\//, ''));
          return { sheet: sheetNameOfPart[sp] || null, tables: out };
        });
      });
      return Promise.all(jobs).then(function (perSheet) {
        var want = [];
        perSheet.forEach(function (ps) { ps.tables.forEach(function (tp) { want.push({ part: tp, sheet: ps.sheet }); }); });
        return Promise.all(want.map(function (w) {
          if (!zip.has(w.part)) return Promise.resolve(null);
          return (isXlsb ? zip.bytes(w.part).then(function (b) { return readTableBin(b); })
            : zip.text(w.part).then(function (t) { return readTableXml(t); }))
            .then(function (def) { if (def) def.sheet = w.sheet; return def; })
            .catch(function () { return null; });
        })).then(function (defs) {
          return defs.filter(function (d) { return !!d && !!d.sheet; });
        });
      });
    });
  }

  /**
   * ★表の名前での参照を A1 範囲に直す★
   *   bytes … 元のファイルのバイト列
   *   kind  … 'xlsb' / 'xlsm' / 'xlsx'
   *   wb    … SheetJS が読んだブック（式の文字と シートの並びを借りる）
   *   Zip   … ZipSurgeon
   * 返り値 { ok, why, fixes:{ 'シート名|r,c': '=直した式' }, stats }
   */
  function resolve(bytes, kind, wb, Zip) {
    var stats = { tables: 0, cells: 0, fixed: 0, refused: 0, skipped: 0 };
    if (!Zip || !wb || !wb.SheetNames) return Promise.resolve({ ok: false, why: '道具が足りない', fixes: {}, stats: stats });
    var isXlsb = (kind === 'xlsb');
    var zip;
    try { zip = Zip.read(bytes); } catch (e) { return Promise.resolve({ ok: false, why: 'zipを開けない', fixes: {}, stats: stats }); }

    return loadTables(zip, isXlsb).then(function (tables) {
      stats.tables = tables.length;
      if (!tables.length) return { ok: true, why: '表が無い', fixes: {}, stats: stats };

      var byId = {}, dupe = false;
      tables.forEach(function (t) { if (byId[t.id]) dupe = true; byId[t.id] = t; });
      if (dupe && isXlsb) return { ok: false, why: '表のidが重なっている', fixes: {}, stats: stats };

      var fixes = {};
      if (!isXlsb) {
        /* ── .xlsx / .xlsm：SheetJS の文字が正しいので、文字を直す ── */
        wb.SheetNames.forEach(function (sn) {
          var ws = wb.Sheets[sn]; if (!ws) return;
          var own = null;
          for (var t = 0; t < tables.length; t++) if (tables[t].sheet === sn) { own = tables[t]; break; }
          Object.keys(ws).forEach(function (a) {
            if (a.charAt(0) === '!') return;
            var f = ws[a].f;
            if (typeof f !== 'string' || f.indexOf('[') < 0) return;
            var rc = decodeAddr(a); if (!rc) return;
            var ownHere = null;
            for (var q = 0; q < tables.length; q++) {
              var d = tables[q];
              if (d.sheet === sn && rc.r >= d.rw1 && rc.r <= d.rw2 && rc.c >= d.cl1 && rc.c <= d.cl2) { ownHere = d; break; }
            }
            stats.cells++;
            var res = rewriteTextFormula('=' + f, tables, { row: rc.r, sheet: sn }, ownHere || own);
            if (res === null) { stats.refused++; return; }
            fixes[sn + '|' + rc.r + ',' + rc.c] = res;
            stats.fixed++;
          });
        });
        return { ok: true, why: '', fixes: fixes, stats: stats };
      }

      /* ── .xlsb：元のバイト列から PtgList を拾って入れ替える ── */
      var names = zip.names().filter(function (n) { return /^xl\/worksheets\/sheet\d+\.bin$/.test(n); });
      var sheetOf = {};
      tables.forEach(function () {});
      /* シート部品 → シート名 を もう一度作る（loadTables と同じ道で作る） */
      return partToSheetName(zip, true).then(function (map) {
        var chain = Promise.resolve();
        names.forEach(function (pn) {
          chain = chain.then(function () {
            var sn = map[pn];
            if (!sn || !wb.Sheets[sn]) { stats.skipped++; return; }
            return zip.bytes(pn).then(function (b) {
              var recs = walkRecords(b);
              if (!recs) { stats.skipped++; return; }
              var rw = 0;
              for (var i = 0; i < recs.length; i++) {
                var r = recs[i];
                if (r.id === 0) { rw = u32(b, r.start); continue; }        // BrtRowHdr
                if (r.id !== 8 && r.id !== 9 && r.id !== 10 && r.id !== 11) continue;
                var rg = rgceRange(r.id, b, r.start, r.len);
                if (!rg) { stats.refused++; continue; }
                var lists = findPtgLists(b, rg);
                if (!lists.length) continue;
                stats.cells++;
                var c = u32(b, r.start);
                var addr = colName(c) + (rw + 1);
                var cell = wb.Sheets[sn][addr];
                var f = cell && typeof cell.f === 'string' ? '=' + cell.f : null;
                if (!f) { stats.refused++; continue; }
                var res = rewriteXlsbFormula(f, lists, byId, { row: rw, sheet: sn });
                if (res === null) { stats.refused++; continue; }
                fixes[sn + '|' + rw + ',' + c] = res;
                stats.fixed++;
              }
            }).catch(function () { stats.skipped++; });
          });
        });
        return chain.then(function () { return { ok: true, why: '', fixes: fixes, stats: stats }; });
      });
    }).catch(function (e) {
      return { ok: false, why: (e && e.message) || String(e), fixes: {}, stats: stats };
    });
  }

  /** シート部品 → シート名 */
  function partToSheetName(zip, isXlsb) {
    var wbPart = isXlsb ? 'xl/workbook.bin' : 'xl/workbook.xml';
    var wbRels = isXlsb ? 'xl/_rels/workbook.bin.rels' : 'xl/_rels/workbook.xml.rels';
    return Promise.all([
      zip.has(wbPart) ? zip.bytes(wbPart) : Promise.resolve(null),
      zip.has(wbRels) ? zip.text(wbRels) : Promise.resolve(''),
    ]).then(function (r) {
      var b = r[0], relsText = r[1] || '', out = {};
      if (!b) return out;
      var relTarget = {}, re = /Id="([^"]+)"[^>]*Target="([^"]+)"/g, m;
      while ((m = re.exec(relsText))) relTarget[m[1]] = m[2];
      var recs = walkRecords(b);
      if (!recs) return out;
      for (var i = 0; i < recs.length; i++) {
        if (recs[i].id !== 156) continue;
        var s = recs[i].start, p = 8, c1 = u32(b, s + p), rel = null;
        if (c1 === 0xFFFFFFFF) p += 4;
        else { rel = ''; for (var k = 0; k < c1; k++) rel += String.fromCharCode(u16(b, s + p + 4 + k * 2)); p += 4 + c1 * 2; }
        var c2 = u32(b, s + p), nm = '';
        for (var k2 = 0; k2 < c2; k2++) nm += String.fromCharCode(u16(b, s + p + 4 + k2 * 2));
        var tgt = rel && relTarget[rel];
        if (tgt) out['xl/' + String(tgt).replace(/^\/?xl\//, '').replace(/^\.\.\//, '')] = nm;
      }
      return out;
    });
  }

  function decodeAddr(a) {
    var m = /^([A-Z]+)(\d+)$/.exec(a);
    if (!m) return null;
    var c = 0;
    for (var i = 0; i < m[1].length; i++) c = c * 26 + (m[1].charCodeAt(i) - 64);
    return { r: parseInt(m[2], 10) - 1, c: c - 1 };
  }

  return {
    resolve: resolve,
    /* ★中を1つずつ試せるように出す（見張りが「入口だけ」を見て緑にならないように）★ */
    _walkRecords: walkRecords,
    _readTableBin: readTableBin,
    _readTableXml: readTableXml,
    _toRange: toRange,
    _rgceRange: rgceRange,
    _findPtgLists: findPtgLists,
    _rewriteXlsbFormula: rewriteXlsbFormula,
    _rewriteTextFormula: rewriteTextFormula,
    _parseRefBody: parseRefBody,
    _colName: colName,
    _partToSheetName: partToSheetName,
    _loadTables: loadTables,
  };
}));
