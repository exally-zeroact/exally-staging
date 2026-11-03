/* xlsb-edit.js — ★.xlsb（BIFF12）の値だけを直す★
 *
 * 形: [記録の番号(可変長1〜2B)][長さ(可変長1〜4B)][中身] の繰り返し
 * ★終端ぴったりで終わらなければ、書き換えを中止して「直せません」と出す★（壊すより断る）
 *
 * 実測（司さんの代行計算表2026・2026-08-09）:
 *   16部品を最後まで歩けた（ズレ0・失敗0）／総記録 43,917／記録の種類 80種
 *   ★数式セル BrtFmlaNum 12,191 + BrtFmlaString 2,866★＝この形式は数式が主役。
 *   だから ★キャッシュ値を自分の計算結果で埋める★のが唯一の担保になる
 *   （.xlsb は再計算しない読み手が多い）。
 *
 * ★binaryIndex について★
 *   xl/worksheets/binaryIndex#.bin は「何バイト目に何がある」の索引。
 *   記録の長さが変わると古い位置を指したままになるので ★消す★。
 *   ただし部品を消すだけでは足りない。★参照が2か所ある★:
 *     ・xl/worksheets/_rels/sheetN.bin.rels の xlBinaryIndex の Relationship
 *     ・[Content_Types].xml の Override（シートの数だけ）
 *   両方を外す。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XlsbEdit = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var R = {
    ROW: 0, BLANK: 1, RK: 2, ERROR: 3, BOOL: 4, REAL: 5, ST: 6, ISST: 7,
    FMLA_STRING: 8, FMLA_NUM: 9, FMLA_BOOL: 10, FMLA_ERROR: 11,
  };
  var VALUE_IDS = [R.BLANK, R.RK, R.ERROR, R.BOOL, R.REAL, R.ST, R.ISST];
  var FORMULA_IDS = [R.FMLA_STRING, R.FMLA_NUM, R.FMLA_BOOL, R.FMLA_ERROR];

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
  function varBytes(v) {
    var out = [];
    do { var b = v & 0x7f; v >>>= 7; if (v) b |= 0x80; out.push(b); } while (v);
    return out;
  }

  /* ★部品ごとの「先頭と末尾の記録」★（実物14シート＋workbook＋sharedStrings で実測・2026-08-09）
   *   終端ぴったりだけでは足りない。★1バイトずらしても終端は合ってしまった★ので、
   *   先頭と末尾も見る。ずらすと先頭が 129 でなくなるので、これで弾ける。 */
  var SHAPE = {
    sheet: { first: 129, last: 130 },
    workbook: { first: 131, last: 132 },
    sharedStrings: { first: 159, last: 160 },
  };
  function shapeOf(partName) {
    if (/worksheets\/sheet\d+\.bin$/.test(partName)) return SHAPE.sheet;
    if (/workbook\.bin$/.test(partName)) return SHAPE.workbook;
    if (/sharedStrings\.bin$/.test(partName)) return SHAPE.sharedStrings;
    return null;                                  // 知らない部品は形を見ない（歩けるかだけ見る）
  }

  /** ★記録を歩く。終端がぴったりでなければ ok:false（＝直させない）★
   *  expect を渡すと、先頭と末尾の記録番号も確かめる（ズレ検出の本命） */
  function parse(buf, expect) {
    var recs = [], p = 0;
    while (p < buf.length) {
      var id = readVar(buf, p, 2);
      if (!id) return { ok: false, why: '番号を読めない（位置 ' + p + '）', recs: recs };
      var q = p + id.bytes;
      var len = readVar(buf, q, 4);
      if (!len) return { ok: false, why: '長さを読めない（位置 ' + q + '）', recs: recs };
      q += len.bytes;
      if (q + len.value > buf.length) {
        return { ok: false, why: '中身がファイルの外へ出る（位置 ' + q + '・長さ ' + len.value + '）', recs: recs };
      }
      recs.push({ id: id.value, data: buf.subarray(q, q + len.value) });
      p = q + len.value;
    }
    if (p !== buf.length) return { ok: false, why: '終端がズレた（' + p + ' / ' + buf.length + '）', recs: recs };
    if (!recs.length) return { ok: false, why: '記録が1つも無い', recs: recs };
    if (expect) {
      var f = recs[0].id, l = recs[recs.length - 1].id;
      if (f !== expect.first) return { ok: false, why: '先頭の記録が違う（' + f + ' / ' + expect.first + 'のはず）', recs: recs };
      if (l !== expect.last) return { ok: false, why: '末尾の記録が違う（' + l + ' / ' + expect.last + 'のはず）', recs: recs };
    }
    return { ok: true, why: '', recs: recs };
  }

  /** 歩いた記録を書き戻す（長さは記録ごとに数え直す＝長さが変わってよい） */
  function build(recs) {
    var total = 0, i;
    for (i = 0; i < recs.length; i++) {
      total += varBytes(recs[i].id).length + varBytes(recs[i].data.length).length + recs[i].data.length;
    }
    var out = new Uint8Array(total), at = 0;
    for (i = 0; i < recs.length; i++) {
      var idB = varBytes(recs[i].id), lnB = varBytes(recs[i].data.length);
      out.set(idB, at); at += idB.length;
      out.set(lnB, at); at += lnB.length;
      out.set(recs[i].data, at); at += recs[i].data.length;
    }
    return out;
  }

  function dv(u8) { return new DataView(u8.buffer, u8.byteOffset, u8.byteLength); }

  /** セルの記録から 列番号を読む（先頭4バイト） */
  function colOf(rec) { return dv(rec.data).getUint32(0, true); }
  /** 行の記録から 行番号を読む */
  function rowOf(rec) { return dv(rec.data).getUint32(0, true); }

  /** 記録の一覧に「行・列」を付けて返す（探しやすくする） */
  function locate(recs) {
    var row = -1, map = {};
    for (var i = 0; i < recs.length; i++) {
      var r = recs[i];
      if (r.id === R.ROW) { row = rowOf(r); continue; }
      if (VALUE_IDS.indexOf(r.id) < 0 && FORMULA_IDS.indexOf(r.id) < 0) continue;
      map[row + ',' + colOf(r)] = i;
    }
    return map;
  }

  /** 数値を入れ直す。
   *  ★数式セル(BrtFmlaNum)は キャッシュの8バイトを その場で書き換える（長さが変わらない）★
   *  値のセルは BrtCellReal に作り直す（長さが変わってよい） */
  function setNumber(recs, idx, num) {
    var r = recs[idx];
    if (r.id === R.FMLA_NUM) {
      var d = r.data.slice();                    // 元を壊さない
      dv(d).setFloat64(8, num, true);            // Cell(8) の直後が答えのキャッシュ
      recs[idx] = { id: r.id, data: d };
      return 'キャッシュを書き換えた（長さ不変）';
    }
    if (FORMULA_IDS.indexOf(r.id) >= 0) {
      /* ══ ★★形が 変わる 式の マスは 「触らない」だけに します★★ ══（2026-09-25）
           ★★何が 起きて いたか（経営者1 が お客さんの 道で 押しました）★★
             ★お客さんが 入力値を 1つ 直して 「書き出す」を 押すと 出ません★
             帯 ･･･「書き出しませんでした／この数式セルの答えの形は まだ直せません（記録 8）」
             ⇒★本番と 同じ 木（15cc377）でも 1字も 違わず 出ません＝元から です★
             ⇒★★＝「直して 保存する」が 出来ません★★（司さんの 決め ア に 直に 当たる）
           ★★因★★
             記録 8 ＝ `BrtFmlaString`（★答えが 字の 式★）。
             ★数を 入れると 記録の 長さが 変わる★ので 書き換えられません。
             ⇒★前は ここで 投げて いました★ ⇒ ★1マスの 為に 本 1冊が 出ません★
           ★★だから 投げずに 「触らなかった」と 数えて 返します★★
             ＝★お客さんが 打った マスは ちゃんと 書かれます★
             ＝★触らなかった 式の 答えは 古いまま 残ります★
             ⇒★★だから 呼ぶ 側が 「開いたら 全部 計算しろ」の 印を 立てます★★
               （`触れなかった` を 返す＝呼ぶ 側が 数えて 印を 立てる）
               ⇒★実Excel が 開いた 時に 計算し直す ので 答えは 合います★
           ★当て推量で 記録を 作り直しては いません★（形を 変えるのは 別の 話）
           ★見張り★ tests/naoshite-hozon-webkit.mjs */
      return { 触れなかった: true, 記録: r.id };
    }
    var cell = r.data.subarray(0, 8);            // 列と書式はそのまま持っていく
    var nd = new Uint8Array(16);
    nd.set(cell, 0);
    dv(nd).setFloat64(8, num, true);
    recs[idx] = { id: R.REAL, data: nd };
    return '値のセルを実数にした';
  }

  /* ══ ★★元の 本に 記録が 無い マスに 数を 足す★★ ══（2026-09-26）
       ★★何が 起きて いたか（経営者1 が お客さんの 道で 押しました）★★
         ★空いて いる マスに 数を 打って 「書き出す」を 押すと 1冊も 出ません★
         帯 ･･･「このファイルは直せません（給料1：★そのセルが元のファイルにありません: 1,8★）」
         ⇒★本番と 同じ 木でも 同じ＝元から です★
         ⇒司さんの 決め ア「★全部 保存しろや★」に 直に 当たります
       ★★当て推量で 記録を 作りません★★
         ・行の 頭（25バイト）は ★元の 板の 行の 頭を 写して 行番号だけ 書き換えます★
           ＝残り 21バイト（高さ・書式 等）を ★自分で 作りません★
           ＝★`xlsb板を作る` が 前から 同じ 手を 使って います★（★前例に 合わせる★）
         ・数の マス（12バイト）＝列(4) ＋ 組の番号(4) ＋ RK(4)
           ＝組の 番号は 0（★飾り 無し★）＝★`xlsb板を作る` と 同じ 決め★
         ・★行の 頭が 1本も 無い 板では 足しません★（★写す 見本が 無い★）
       ★★並びを 壊しません★★
         ・行は ★行番号の 順★／マスは ★列の 順★ に 入れます
         ・★使った 範囲（記録148）も 広げます★（★広げないと Excel が 見ません★）
       ★見張り★ tests/akimasu-ni-utsu-webkit.mjs */
  function マスを足す(recs, r, c, num) {
    /* ★写す 見本（行の 頭）を 探します★ */
    var 見本 = null;
    for (var i = 0; i < recs.length; i++) {
      if (recs[i].id === R.ROW && recs[i].data && recs[i].data.length === 25) { 見本 = recs[i].data; break; }
    }
    if (!見本) return { ok: false, why: '行の頭の見本が在りません（作りません）' };

    var 中身の始め = -1, 中身の終り = -1;
    for (i = 0; i < recs.length; i++) {
      if (recs[i].id === 145) 中身の始め = i;
      if (recs[i].id === 146) { 中身の終り = i; break; }
    }
    if (中身の始め < 0 || 中身の終り < 0) return { ok: false, why: '中身の 始め/終りが 在りません' };

    /* ★その 行の 頭を 探します★／無ければ ★行番号の 順に 入れる 所★ を 決めます */
    var 行の所 = -1, 入れる所 = 中身の終り, 今の行 = -1;
    for (i = 中身の始め + 1; i < 中身の終り; i++) {
      if (recs[i].id !== R.ROW) continue;
      var rr = rowOf(recs[i]);
      if (rr === r) { 行の所 = i; break; }
      if (rr > r) { 入れる所 = i; 今の行 = rr; break; }
    }

    var 新しい行 = false;
    if (行の所 < 0) {
      var 頭 = new Uint8Array(25);
      頭.set(見本, 0);
      _4(頭, 0, r);
      recs.splice(入れる所, 0, { id: R.ROW, data: 頭 });
      行の所 = 入れる所;
      新しい行 = true;
      中身の終り++;
    }

    /* ★その 行の 中で 列の 順に 入れる 所を 決めます★ */
    var 置く = 行の所 + 1;
    for (i = 行の所 + 1; i < 中身の終り; i++) {
      if (recs[i].id === R.ROW) break;
      var cc = colOf(recs[i]);
      if (cc === c) return { ok: false, why: 'そのマスは もう 在ります（' + r + ',' + c + '）' };
      if (cc > c) break;
      置く = i + 1;
    }
    var m = new Uint8Array(12);
    var v = new DataView(m.buffer, m.byteOffset, 12);
    v.setUint32(0, c, true);
    v.setUint32(4, 0, true);      /* ★組の 番号 0＝飾り 無し★ */
    v.setUint32(8, RKにする(num), true);
    recs.splice(置く, 0, { id: XLSB_R.数のマス, data: m });

    /* ★使った 範囲を 広げます★（★広げないと Excel が 見ません★） */
    for (i = 0; i < recs.length; i++) {
      if (recs[i].id !== XLSB_R.範囲) continue;
      if (!recs[i].data || recs[i].data.length < 16) break;
      var d = recs[i].data.slice();
      var dv2 = new DataView(d.buffer, d.byteOffset, d.byteLength);
      var r1 = dv2.getUint32(0, true), r2 = dv2.getUint32(4, true);
      var c1 = dv2.getUint32(8, true), c2 = dv2.getUint32(12, true);
      if (r < r1) dv2.setUint32(0, r, true);
      if (r > r2) dv2.setUint32(4, r, true);
      if (c < c1) dv2.setUint32(8, c, true);
      if (c > c2) dv2.setUint32(12, c, true);
      recs[i] = { id: recs[i].id, data: d };
      break;
    }
    return { ok: true, 新しい行: 新しい行 };
  }

  /** シート1枚の中身（Uint8Array）を受け取り、値を入れ直して返す */
  function editSheet(bin, cells) {
    var p = parse(bin, SHAPE.sheet);
    if (!p.ok) return { ok: false, why: p.why };
    var map = locate(p.recs);
    var done = [], miss = [];
    var 触れなかった = [];
    var 足した = [];
    Object.keys(cells).forEach(function (key) {         // key = "行,列"（0始まり）
      var i = map[key];
      if (i === undefined) {
        /* ══ ★★元の 本に 無い マスは 足します★★ ══（2026-09-26）
             ＝★前は ここで 本 1冊 まるごと 断って いました★
             ＝実測（経営者1・お客さんの 道）･･･ ★空きマスに 打つと 1冊も 出ない★
             ⇒★足せない 時だけ 断ります★（行の 頭の 見本が 無い 等） */
        var 番 = String(key).split(',');
        var 足 = マスを足す(p.recs, parseInt(番[0], 10), parseInt(番[1], 10), cells[key]);
        if (!足.ok) { miss.push(key + '（' + 足.why + '）'); return; }
        足した.push(key + (足.新しい行 ? ':行も足した' : ':マスを足した'));
        return;
      }
      var 出 = setNumber(p.recs, i, cells[key]);
      /* ★形が 変わる 式の マスは 触らずに 数えます★（上の 断りを 見て ください） */
      if (出 && 出.触れなかった) { 触れなかった.push(key + ':記録' + 出.記録); return; }
      done.push(key + ':' + 出);
    });
    if (miss.length) return { ok: false, why: 'そのマスを 足せません: ' + miss.join(' ') };
    /* ★触れなかった 数も 返します★＝★呼ぶ 側が 「開いたら 全部 計算しろ」の 印を 立てる★ */
    return { ok: true, bytes: build(p.recs), done: done, count: p.recs.length,
      触れなかった: 触れなかった, 足した: 足した };
  }

  /** binaryIndex を外す（★部品・rels・[Content_Types].xml の3か所★） */
  function dropBinaryIndex(zip, contentTypesXml, relsBySheet) {
    var dropped = [], names = zip.names();
    names.forEach(function (n) {
      if (/^xl\/worksheets\/binaryIndex\d+\.bin$/.test(n)) { zip.remove(n); dropped.push(n); }
    });
    // [Content_Types].xml の Override を外す
    var ct = contentTypesXml.replace(/<Override[^>]*binaryIndex\d+\.bin[^>]*\/>/g, '');
    // 各シートの rels から xlBinaryIndex の参照を外す
    var rels = {};
    Object.keys(relsBySheet).forEach(function (k) {
      rels[k] = relsBySheet[k].replace(/<Relationship[^>]*xlBinaryIndex[^>]*\/>/g, '');
    });
    return { dropped: dropped, contentTypes: ct, rels: rels };
  }

  /* ══ ★★板の 名前 ⇒ 部品名（`.xlsb`）★★ ══（2026-09-21）
       ★★なぜ 要るか★★
         この ファイルの `saveXlsb` は こう 書いて あります:
           「workbook.bin を 読まずに、開いた時の 並びで 対応させる」
         ＝★`sheet1.bin` が 1枚目★ と 当てて います。
         ⇒★合わないと 判子や 値が 別の 板に 行きます★（★消えるのでは なく ずれる★）
       ★★2026-09-21 に 実物で 割りました★★（経営者1 が 材料を 2本 作りました）
         ①`exally-tameshi-xlsb2.xlsb` ... 1枚目 Ita1 ／ 2枚目 Ita2
         ②`exally-tameshi-xlsb3.xlsb` ... ★板の 並びを 入れ替えた 物★（1枚目 Ita2）
         ⇒②でも ★Excel が 番号を 並びに 付け直して いました★
           （板1=Ita2 が `rId1` ⇒ `sheet1.bin`）
         ⇒★★つまり ①②では 「番号当て」でも 同じ 答えに なります★★
         ⇒★★だから この 2本では 番号当ての 良し悪しは 割れません★★
           （経営者1 が そう 書いて くれました＝★言い過ぎない★）
       ★★それでも 番号当てを やめます★★
         ＝★「たまたま 合って いる」と 「引いて いる」は 別★
         ＝★他の 道具（LibreOffice・古い Excel・板を 消した 後）は 未測定★
         ⇒★`workbook.bin` の 並び ⇒ rId ⇒ rels ⇒ 部品名★ を 引きます。
       ★★記録の 形★★（2026-09-21 実測・板 2枚の `.xlsb` 2本）
         `BrtBundleSh` ＝ ★記録 156★（板 1枚に つき 1本）
           hsState  ... 4バイト（隠れて いるか）
           iTabID   ... 4バイト
           strRelID ... ★4バイトの 字数 ＋ UTF-16LE★（`0xFFFFFFFF` は 無し）
           strName  ... 同じ 形
         ⇒★並びは 記録の 並び そのもの★（`iTabID` の 順では ありません）
           ＝②で `iTabID` は 2,1 の 順に 出ました（★並べ替えに 使えません★）
       ★★読めない 時は null★★＝★呼ぶ 側が 前の やり方に 戻れます★ */
  var BrtBundleSh = 156;
  /** ★`4バイトの 字数 ＋ UTF-16LE` を 読む★（無しは null） */
  function 字を読む(d, p) {
    if (p + 4 > d.byteLength) return null;
    var n = d.getUint32(p, true);
    if (n === 0xFFFFFFFF) return { 字: null, 次: p + 4 };
    if (n > 32767 || p + 4 + n * 2 > d.byteLength) return null;
    var s = '';
    for (var i = 0; i < n; i++) s += String.fromCharCode(d.getUint16(p + 4 + i * 2, true));
    return { 字: s, 次: p + 4 + n * 2 };
  }
  /** ★★口★★ ＝ `workbook.bin` と `xl/_rels/workbook.bin.rels` から
   *    `[{ 名, rId, 部品 }]` を ★板の 並びの 順に★ 返す（★読めなければ null★） */
  function 板たち(workbookBin, relsXml) {
    if (!workbookBin || !relsXml) return null;
    var u8 = (workbookBin instanceof Uint8Array) ? workbookBin : new Uint8Array(workbookBin);
    var r = parse(u8, SHAPE.workbook);
    if (!r.ok) return null;
    /* ★rels は XML の まま 残ります★（`.xlsb` でも）＝字で 引けます */
    var 行き先 = {};
    var 字 = String(relsXml), i = 0;
    while (true) {
      var a = 字.indexOf('<Relationship', i);
      if (a < 0) break;
      var g = 字.indexOf('>', a);
      if (g < 0) break;
      var 札 = 字.slice(a, g);
      var id = 抜く(札, 'Id'), t = 抜く(札, 'Target');
      if (id && t) 行き先[id] = t;
      i = g + 1;
    }
    var 出 = [];
    for (var k = 0; k < r.recs.length; k++) {
      if (r.recs[k].id !== BrtBundleSh) continue;
      var b = r.recs[k].data;
      var d = new DataView(b.buffer, b.byteOffset, b.byteLength);
      var rid = 字を読む(d, 8);
      if (!rid) return null;
      var nm = 字を読む(d, rid.次);
      if (!nm || nm.字 === null) return null;
      var t2 = rid.字 ? 行き先[rid.字] : null;
      if (!t2) return null;                       /* ★引けない 物が 1つでも 在れば 全部 断る★ */
      出.push({ 名: nm.字, rId: rid.字,
        部品: (t2.charAt(0) === '/') ? t2.slice(1) : ('xl/' + t2.replace(/^\.\//, '')) });
    }
    return 出.length ? 出 : null;
  }
  function 抜く(札, 名) {
    var k = ' ' + 名 + '="';
    var a = 札.indexOf(k);
    if (a < 0) return null;
    var b = 札.indexOf('"', a + k.length);
    return b < 0 ? null : 札.slice(a + k.length, b);
  }
  /* ══ ★★`.xlsb` の 板を 作る★★ ══（2026-09-22）
       ★★なぜ 要るか★★
         司さん「★全部 保存しろや★」（ア）＝★司さんの 実物は `.xlsb`★
         `.xlsx` の 側は 2026-09-21 に 直しました（`lib/xlsx-edit.js` の `板を足す`）。
         ⇒`.xlsb` は ★別の 道★（`workbook.bin` も `sheetN.bin` も 2進）

       ★★実物で 数えた 事★★（`tests/fixtures/kazari-hiraku3.xlsb`・2026-09-22）
         ★板の 骨★（記録の 並び）
           129 ... 板の 始め（中身 0バイト）
           147 ... 板の 決め（23バイト）
           148 ... ★使った 範囲★（16バイト）＝行の 初/終・列の 初/終 の 4つ
           133 137 152 138 134 ... 見え方（省いても 歩けます）
           145 ... ★中身の 始め★
             0 ... ★行の 頭★（25バイト・頭の 4バイトが 行番号）
             2 ... ★数の マス★（★★12バイト★★）
                 ＝列(4) + ★組の番号★ + ★RK(4)★
                 ★★組の 番号が 24ビットか 32ビットかは この 材料では 分かれません★★
                   ＝8バイト目（byte7）が ★どの マスも 0★ だから
                   ＝★どちらでも 今の 実物では 実害は 在りません★
                   ＝経営者1 の 断り（2026-09-22）を そのまま 写して います
                   ⇒★組の 番号が 大きい 本で 測れば 分かれます★（未測定）
                 ★★`double` では ありません★★＝12バイトに 入りません
                 ＝経営者1 の 紙（85）は 「8-11 数(double)」と 書いて いますが
                   ★そこは 4バイトです★（2026-09-22 実物で 数え直しました）
                 ★★RK＝縮めた 数★★（下位 2ビットが 印）
                   ビット0 ... 立つと ★100で 割る★
                   ビット1 ... 立つと ★整数★（上位30ビットを 2つ ずらす）
                              立たないと ★実数の 上位 4バイト★（下位は 0）
                 ★実物で 合わせました★ ... 3 / 1 / 0.25 が そのまま 出ます
                 ★double の マス（記録 5）は この 材料に 1つも 在りません★
             7 ... ★文字の マス★（12バイト）＝`sharedStrings.bin` の 番号を 指す
           146 ... 中身の 終り
           130 ... 板の 終り
         ★`XlsbEdit.SHAPE.sheet` が 129/130 を 見て います★（前から 在る 門）

       ★★文字は 書きません★★（★測って いない から★）
         `.xlsb` の 文字の マス（記録 7）は ★`sharedStrings.bin` の 番号★を 指します。
         ⇒足すには ★共有文字列の 側も 直す★事に なります
         ⇒★元の 板が 指す 番号が ずれる 恐れ★が 在ります
         ⇒`.xlsx` では `inlineStr` で 逃げましたが ★`.xlsb` に 同じ 物が 在るかは 未測定★
         ⇒★だから 今は 数だけ 書きます★
         ⇒★字の マスは 「入りません」と 言い続けます★（`lib/hairanai.js`）

       ★★測って いない 事★★
         ・★実Excel が この 板を どう 開くかは ここでは 測れません★（COM は 経営者1 の 持ち場）
         ・見え方の 記録（133/137/152/138/134）を 省いて 良いかは ★未測定★
           ⇒★省かずに 元の 板から 写します★（★消すより 残す★）
         ・`xl/metadata.bin` は ★1バイトも 触りません★（★触ると 溢れが 壊れます★・経営者1 の 実測） */
  var XLSB_R = {
    板の始め: 129, 板の終り: 130, 決め: 147, 範囲: 148,
    中身の始め: 145, 中身の終り: 146, 行の頭: 0, 数のマス: 2,
  };

  /* ══ ★★RK＝縮めた 数★★ ══（2026-09-22 実物で 数えました）
       下位 2ビットが 印
         ビット0 ... 立つと ★100で 割る★
         ビット1 ... 立つと ★整数★（上位30ビットを 2つ ずらす）
                    立たないと ★実数の 上位 4バイト★（下位 4バイトは 0）
       ★★入らない 数は 在ります★★
         ＝下位 4バイトが 0 で ない 実数（例 0.1）
         ⇒★その時は 100倍が 整数に なるか を 見ます★
         ⇒★それでも 入らなければ 投げます★＝★黙って 丸めません★
         ⇒（`double` の マス＝記録 5 は ★まだ 書いて いません★） */
  function RKにする(数) {
    /* ★①整数（30ビットに 入る）★ */
    if (Number.isInteger(数) && 数 >= -(1 << 29) && 数 < (1 << 29)) {
      return (((数 << 2) | 2) >>> 0);
    }
    /* ★②100倍が 整数（30ビットに 入る）★ */
    var 百 = Math.round(数 * 100);
    if (Math.abs(百 / 100 - 数) < 1e-12 && 百 >= -(1 << 29) && 百 < (1 << 29)) {
      return (((百 << 2) | 3) >>> 0);
    }
    /* ★③実数の 下位 4バイトが 0 なら そのまま★ */
    var bb = new ArrayBuffer(8);
    var dv = new DataView(bb);
    dv.setFloat64(0, 数, true);
    if (dv.getUint32(0, true) === 0) {
      var 上 = dv.getUint32(4, true);
      if ((上 & 3) === 0) return 上 >>> 0;
    }
    throw new Error('この数は .xlsb の縮めた数に入りません（' + 数 + '）');
  }

  /** ★4バイトの 数を 入れる★（下位から） */
  function _4(out, at, v) {
    out[at] = v & 0xff; out[at + 1] = (v >>> 8) & 0xff;
    out[at + 2] = (v >>> 16) & 0xff; out[at + 3] = (v >>> 24) & 0xff;
  }

  /** ★`.xlsb` の 板を 1枚 作る★（★数だけ★・字は 書きません）
   *  @param 見本 元の 板の 記録（`parse()` の `recs`）＝★見え方を そのまま 写す為★
   *  @param cells `{ 'A1': { v, t } }`
   *  戻り ... `Uint8Array`（`build()` で 包みに 入れられる 形） */
  function xlsb板を作る(見本, cells) {
    var 出 = [];
    var 写す = function (id) {
      for (var i = 0; 見本 && i < 見本.length; i++) {
        if (見本[i].id === id) { 出.push({ id: id, data: 見本[i].data }); return true; }
      }
      return false;
    };

    出.push({ id: XLSB_R.板の始め, data: new Uint8Array(0) });
    写す(XLSB_R.決め);                       /* ★無くても 歩けます★＝在れば 写す */

    /* ★★使った 範囲★★＝行の 初/終・列の 初/終（★0 から 数えます★） */
    var 行ら = [], 列ら = [];
    var 表 = {};
    Object.keys(cells || {}).forEach(function (a) {
      var m = /^([A-Z]+)(\d+)$/.exec(a);
      if (!m) return;
      var c = 0, s = m[1];
      for (var i = 0; i < s.length; i++) c = c * 26 + (s.charCodeAt(i) - 64);
      c -= 1;
      var r = parseInt(m[2], 10) - 1;
      var v = cells[a] && cells[a].v;
      if (typeof v !== 'number') return;     /* ★数だけ★（上の 注） */
      行ら.push(r); 列ら.push(c);
      (表[r] = 表[r] || []).push({ 列: c, 数: v });
    });
    var 範 = new Uint8Array(16);
    _4(範, 0, 行ら.length ? Math.min.apply(null, 行ら) : 0);
    _4(範, 4, 行ら.length ? Math.max.apply(null, 行ら) : 0);
    _4(範, 8, 列ら.length ? Math.min.apply(null, 列ら) : 0);
    _4(範, 12, 列ら.length ? Math.max.apply(null, 列ら) : 0);
    出.push({ id: XLSB_R.範囲, data: 範 });

    /* ★見え方は 元の 板から 写します★（★消すより 残す★・省いて 良いかは 未測定） */
    [133, 137, 152, 138, 134].forEach(写す);

    出.push({ id: XLSB_R.中身の始め, data: new Uint8Array(0) });
    var 行番 = Object.keys(表).map(Number).sort(function (a, b) { return a - b; });
    var 行の見本 = null;
    for (var k = 0; 見本 && k < 見本.length; k++) {
      if (見本[k].id === XLSB_R.行の頭) { 行の見本 = 見本[k].data; break; }
    }
    行番.forEach(function (r) {
      /* ★行の 頭は 25バイト★＝★元の 物を 写して 行番号だけ 書き換えます★
           ＝残り 21バイト（高さ・書式 等）を ★当て推量で 作らない★ */
      var 頭 = new Uint8Array(25);
      if (行の見本 && 行の見本.length === 25) 頭.set(行の見本, 0);
      _4(頭, 0, r);
      出.push({ id: XLSB_R.行の頭, data: 頭 });
      表[r].sort(function (a, b) { return a.列 - b.列; }).forEach(function (c) {
        /* ★数の マスは 12バイト★＝列(4) + 組の番号(4) + ★RK(4)★
             ★組の 番号は 0★＝★飾り 無し★（`.xlsx` の 側と 同じ 決め） */
        var m = new Uint8Array(12);
        var v = new DataView(m.buffer, m.byteOffset, 12);
        v.setUint32(0, c.列, true);
        v.setUint32(4, 0, true);   /* ★組の 番号 0＝飾り 無し★
             ★0 なら 24ビットでも 32ビットでも 同じ 字に なります★ */
        v.setUint32(8, RKにする(c.数), true);
        出.push({ id: XLSB_R.数のマス, data: m });
      });
    });
    出.push({ id: XLSB_R.中身の終り, data: new Uint8Array(0) });
    出.push({ id: XLSB_R.板の終り, data: new Uint8Array(0) });
    return build(出);
  }

  /* ══ ★★`.xlsb` に 板を 足す★★ ══（2026-09-22）
       ★★直す 部品は 4つ★★（`.xlsx` と 同じ 形・★中身は 別★）
         ①`xl/worksheets/sheetN.bin` ... ★2進★（`XlsbEdit.xlsb板を作る`）
         ②`xl/workbook.bin` の `BrtBundleSh`（記録 156）に 1本 足す ... ★2進★
         ③`xl/_rels/workbook.bin.rels` ... ★XML の まま★
         ④`[Content_Types].xml` ... ★XML の まま★
            ★型の 名は `.xlsx` と 違います★（2026-09-22 実物で 数えました）
              `.xlsb` ... `application/vnd.ms-excel.worksheet`
              `.xlsx` ... `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`
       ★★`xl/metadata.bin` は 1バイトも 触りません★★
         ＝★触ると 溢れが 壊れます★（経営者1 の 実測）
       ★★`xl/worksheets/binaryIndex1.bin` は 消します★★
         ＝「何バイト目に 何が 在る」の 索引＝★板が 増えると 嘘に なります★
         ＝`XlsbEdit.dropBinaryIndex` が 前から 在ります（★同じ 形の 前例★）
         ⇒★但し 「消して 良い」は まだ 測って いません★＝経営者1 が 開いて 数えます
       ★★字の マスは 書きません★★＝`sharedStrings.bin` を 触る 事に なる
         ⇒★元の 板が 指す 番号が ずれる 恐れ★
         ⇒`lib/hairanai.js` は `.xlsb` で 今まで 通り「入りません」と 言います */
  function xlsb板を足す(zip, 見本の板, name, cells) {
    var 番 = 1;
    zip.names().forEach(function (n) {
      var m = /^xl\/worksheets\/sheet(\d+)\.bin$/.exec(n);
      if (m) { var v = parseInt(m[1], 10); if (v >= 番) 番 = v + 1; }
    });
    var 部品 = 'xl/worksheets/sheet' + 番 + '.bin';
    return zip.text('xl/_rels/workbook.bin.rels').then(function (rels) {
      /* ★空いて いる rId★＝★数え直さずに 決めると ぶつかります★ */
      var 大 = 0, i = 0, 鍵 = ' Id="rId';
      while (true) {
        var a = rels.indexOf(鍵, i);
        if (a < 0) break;
        var b = rels.indexOf('"', a + 鍵.length);
        if (b < 0) break;
        var v = parseInt(rels.slice(a + 鍵.length, b), 10);
        if (isFinite(v) && v > 大) 大 = v;
        i = b + 1;
      }
      var rId = 'rId' + (大 + 1);

      /* ①板の 中身 */
      zip.add(部品, xlsb板を作る(見本の板, cells));

      /* ③rels */
      zip.replaceText('xl/_rels/workbook.bin.rels', rels.replace('</Relationships>',
        '<Relationship Id="' + rId + '"'
        + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
        + ' Target="worksheets/sheet' + 番 + '.bin"/></Relationships>'));

      return zip.text('[Content_Types].xml').then(function (ct) {
        /* ④型 */
        if (ct.indexOf('</Types>') < 0) throw new Error('[Content_Types].xml が読めません');
        zip.replaceText('[Content_Types].xml', ct.replace('</Types>',
          '<Override PartName="/' + 部品 + '"'
          + ' ContentType="application/vnd.ms-excel.worksheet"/></Types>'));

        /* ②`workbook.bin` に 板を 1本 足す */
        return zip.bytes('xl/workbook.bin').then(function (wbb) {
          var w = parse(wbb instanceof Uint8Array ? wbb : new Uint8Array(wbb), SHAPE.workbook);
          if (!w.ok) throw new Error('workbook.bin を歩けません（' + w.why + '）');
          /* ★最後の `BrtBundleSh` の 後ろへ 入れます★＝★並びが 板の 順★ */
          var 後 = -1, 最大id = 0;
          for (var k = 0; k < w.recs.length; k++) {
            if (w.recs[k].id !== BrtBundleSh) continue;
            後 = k;
            var d = w.recs[k].data;
            var dv = new DataView(d.buffer, d.byteOffset, d.byteLength);
            var t = dv.getUint32(4, true);
            if (t > 最大id) 最大id = t;
          }
          if (後 < 0) throw new Error('workbook.bin に板が1枚も見つかりません');
          w.recs.splice(後 + 1, 0, { id: BrtBundleSh, data: BundleShを作る(最大id + 1, rId, name) });
          zip.replace('xl/workbook.bin', build(w.recs));
          return { 部品: 部品, rId: rId, iTabID: 最大id + 1 };
        });
      });
    });
  }

  /** ★`BrtBundleSh` を 作る★＝hsState(4) + iTabID(4) + strRelID + strName
   *  ★字は 4バイトの 字数 ＋ UTF-16LE★（`板たち()` が 読む 形と 同じ） */
  function BundleShを作る(iTabID, rId, name) {
    var 字 = function (s) {
      var u = new Uint8Array(4 + s.length * 2);
      var v = new DataView(u.buffer);
      v.setUint32(0, s.length, true);
      for (var i = 0; i < s.length; i++) v.setUint16(4 + i * 2, s.charCodeAt(i), true);
      return u;
    };
    var a = 字(String(rId)), b = 字(String(name));
    var out = new Uint8Array(8 + a.length + b.length);
    var dv = new DataView(out.buffer);
    dv.setUint32(0, 0, true);          /* hsState ＝ 0（隠れて いない） */
    dv.setUint32(4, iTabID, true);
    out.set(a, 8);
    out.set(b, 8 + a.length);
    return out;
  }

  /* ══ ★★「開いたら 全部 計算しろ」の 印を 立てる★★ ══（2026-09-25）
       ★★なぜ 要るか★★
         `editSheet` は ★答えが 字の 式（記録 8）を 触りません★。
         ⇒★その 答えは 古いまま 残ります★
         ⇒★★だから 実Excel に 「開いたら 計算し直して」と 頼みます★★
         ⇒★そう すれば 古い 答えは 開いた 時に 直ります★
       ★★番号は 記憶で 書いて いません★★
         `workbook.bin` の ★記録 157（26バイト）★ の ★26バイト目（0から 25）の ビット0★
         ＝経営者1 が ★実Excel で 同じ 中身の 本を 7冊 作り★、私が 読み直して 決めました
         ＝読む 側は `lib/xlsb-jitai.js` の `開いたら全部計算するか`（★同じ 番号★）
       ★記録が 無い 時は 触りません★＝★当て推量で 作らない★
       ★見張り★ tests/naoshite-hozon-webkit.mjs */
  var 計算の決めの記録 = 157;
  var 全部計算のバイト = 25;
  function 全部計算の印を立てる(wbBytes) {
    var u8 = wbBytes instanceof Uint8Array ? wbBytes : new Uint8Array(wbBytes);
    var r = parse(u8, SHAPE.workbook);
    if (!r.ok) return { ok: false, why: r.why };
    for (var i = 0; i < r.recs.length; i++) {
      var rec = r.recs[i];
      if (rec.id !== 計算の決めの記録) continue;
      if (!rec.data || rec.data.length <= 全部計算のバイト) return { ok: false, why: '記録157が 短い' };
      if ((rec.data[全部計算のバイト] & 0x01) !== 0) return { ok: true, bytes: u8, もう立っていた: true };
      var d = rec.data.slice();
      d[全部計算のバイト] = d[全部計算のバイト] | 0x01;
      r.recs[i] = { id: rec.id, data: d };
      return { ok: true, bytes: build(r.recs), 立てた: true };
    }
    return { ok: false, why: '記録157が 在りません' };
  }

  return {
    全部計算の印を立てる: 全部計算の印を立てる,
    R: R, VALUE_IDS: VALUE_IDS, FORMULA_IDS: FORMULA_IDS, SHAPE: SHAPE, shapeOf: shapeOf,
    parse: parse, build: build, locate: locate, editSheet: editSheet,
    dropBinaryIndex: dropBinaryIndex, colOf: colOf, rowOf: rowOf,
    板たち: 板たち, BrtBundleSh: BrtBundleSh,
    xlsb板を作る: xlsb板を作る, XLSB_R: XLSB_R,
    xlsb板を足す: xlsb板を足す, RKにする: RKにする,
  };
});
