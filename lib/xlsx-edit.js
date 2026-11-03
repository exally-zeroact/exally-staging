/* xlsx-edit.js — ★受け取ったブックを「触った所だけ」直す★（.xlsx / .xlsm）
 *
 * ★なぜ作り直さないか（2026-08-08 実測）★
 *   SheetJS で読んで書き直すと 罫線84・判子1・結合7・列幅が★全部消える★。
 *   だから ★元のzipを持ったまま、セルの値だけ書き換えて閉じ直す★。
 *
 * ★守っている決まり（全部 実測で決めた）★
 *   ・シートは ★名前 → r:id → rels の Target★ で引く（`sheet1.xml` は先頭シートではない）
 *   ・文字列は sharedStrings に ★末尾へ足して番号だけ差し替える★
 *     （既にある <si> を書き換えると、同じ言葉のセルが全部変わる。実測で3セル巻き込んだ）
 *   ・★数式セルは <f> を残し、<v>(答えのキャッシュ)を自分の計算結果で埋める★
 *     Excelは開いても再計算しない。fullCalcOnLoad は Excel でしか効かないので、
 *     ★Google/Numbers/PDF変換/メールのプレビューのためには「自分で埋める」しかない★
 *   ・その上で ★fullCalcOnLoad="1" も立てる（保険）★
 *   ・書式(s属性)は触らない。無いセルは★作らずに断る★（壊すより断る）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./zip-surgeon.js'));
  else root.XlsxEdit = factory(root.ZipSurgeon);
})(typeof self !== 'undefined' ? self : this, function (ZipSurgeon) {
  'use strict';

  function esc(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  /* 属性の中に出てくる特殊文字を正規表現用に逃がす */
  function rx(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  /** ブックを開く。★元のバイト列は zip.raw に残る（作業用コピーはこれを保存する）★ */
  function open(bytes) {
    var zip = ZipSurgeon.read(bytes);
    var book = { zip: zip, sheets: [], sst: null, dirty: {} };
    return zip.text('xl/workbook.xml').then(function (wbx) {
      book.workbookXml = wbx;
      return zip.text('xl/_rels/workbook.xml.rels');
    }).then(function (rels) {
      book.relsXml = rels;
      var re = /<sheet\b[^>]*\/>/g, m;
      while ((m = re.exec(book.workbookXml))) {
        var tag = m[0];
        var name = (tag.match(/\bname="([^"]*)"/) || [])[1];
        var rid = (tag.match(/r:id="([^"]*)"/) || [])[1];
        if (!name || !rid) continue;
        var t = book.relsXml.match(new RegExp('<Relationship[^>]*Id="' + rx(rid) + '"[^>]*Target="([^"]+)"'));
        if (!t) continue;
        var target = t[1].replace(/^\//, '').replace(/^xl\//, '');
        book.sheets.push({ name: name, rid: rid, part: 'xl/' + target });
      }
      if (!book.sheets.length) throw new Error('シートが1枚も読み取れません（この形は まだ直せません）');
      return book;
    });
  }

  function sheetOf(book, name) {
    for (var i = 0; i < book.sheets.length; i++) if (book.sheets[i].name === name) return book.sheets[i];
    throw new Error('そのシートがありません: ' + name);
  }

  /* ── 共有文字列：★末尾に足して番号を返す★（既にある物は書き換えない） ── */
  function ensureSst(book) {
    if (book.sst) return Promise.resolve(book.sst);
    var name = 'xl/sharedStrings.xml';
    if (!book.zip.has(name)) {
      book.sst = { xml: null, count: 0, added: [] };   // 共有文字列が無いブックもある
      return Promise.resolve(book.sst);
    }
    return book.zip.text(name).then(function (xml) {
      var n = (xml.match(/<si[\s>]/g) || []).length;
      book.sst = { xml: xml, count: n, added: [] };
      return book.sst;
    });
  }
  function addString(book, s) {
    var sst = book.sst;
    var idx = sst.count + sst.added.length;
    sst.added.push(s);
    return idx;
  }

  /** セル1つを書き換える予約。
   *  spec: { v: 値, t: 'n'|'s'|'b', f: 式が有る場合は触らない }
   *  ★数式セルは <f> を残して <v> だけ入れ替える★ */
  function setCell(sheetXml, addr, spec, book) {
    var re = new RegExp('<c\\b([^>]*\\br="' + rx(addr) + '"[^>]*)(/>|>([\\s\\S]*?)</c>)');
    var m = sheetXml.match(re);
    if (!m) {
      throw new Error('セル ' + addr + ' が元のファイルにありません（新しいセルは まだ作れません）');
    }
    var attrs = m[1];
    var inner = m[3] || '';
    var hasF = /<f[\s>]/.test(inner);
    var fPart = hasF ? (inner.match(/<f[\s\S]*?<\/f>|<f\b[^>]*\/>/) || [''])[0] : '';

    // 型を表す t 属性を作り直す（書式 s= は そのまま残す）
    var keep = attrs.replace(/\s+t="[^"]*"/g, '');
    var t = '', body = '';
    if (spec.t === 's') {
      var idx = addString(book, spec.v == null ? '' : String(spec.v));
      t = ' t="s"'; body = '<v>' + idx + '</v>';
    } else if (spec.t === 'b') {
      t = ' t="b"'; body = '<v>' + (spec.v ? 1 : 0) + '</v>';
    } else {
      t = ''; body = '<v>' + Number(spec.v) + '</v>';   // 数値は t を付けないのが既定
    }
    var out = '<c' + keep + t + '>' + fPart + body + '</c>';
    return sheetXml.slice(0, m.index) + out + sheetXml.slice(m.index + m[0].length);
  }

  /** シート1枚ぶんの書き換えを予約する。cells = { 'A1': {v,t}, ... } */
  function setValues(book, sheetName, cells) {
    var sh = sheetOf(book, sheetName);
    return ensureSst(book).then(function () {
      return book.dirty[sh.part] ? Promise.resolve(book.dirty[sh.part]) : book.zip.text(sh.part);
    }).then(function (xml) {
      Object.keys(cells).forEach(function (addr) {
        xml = setCell(xml, addr, cells[addr], book);
      });
      book.dirty[sh.part] = xml;
      return book;
    });
  }

  /* ══ ★★板を 足す★★ ══（2026-09-21）
       ★★なぜ 要るか★★
         司さん「★全部 保存しろや、断る 理由が なんか あるんか★」（ア）
         `js/book-open.js` の `saveXlsxLike()` は こう 書いて ありました
           `if (opened.sheetNames.indexOf(sh.name) < 0) return;`
         ⇒★うちで 足した 板は 書き出す 先に 入りません★
           （`lib/hairanai.js` が 数えて 言う 物）
         ⇒★「出来ない から」では ありません★＝足りないのは 口 だけ でした。
       ★★直す 部品は 4つ★★（★1つでも 抜けると 実Excel が 修復を 言います★）
         ①`xl/worksheets/sheetN.xml` ... 板の 中身（★足す★）
         ②`xl/workbook.xml` の `<sheets>` ... 1行 足す（名前・`sheetId`・`r:id`）
         ③`xl/_rels/workbook.xml.rels` ... 1行 足す（`rId` ⇒ 部品名）
         ④`[Content_Types].xml` ... 1行 足す（板の 型）
       ★★番号は 空いて いる 物を 使います★★
         `sheetN.xml` の N ／ `rIdN` の N ／ `sheetId`
         ＝どれも ★今 在る 一番 大きい 数 ＋ 1★
         ＝★数え直さずに 決めると ぶつかります★
       ★★`xl/calcChain.xml` は 触りません★★
         ＝★消して 良いかを 測って いません★（★見立てで 消さない★）
         ＝`save()` が `fullCalcOnLoad` を 立てるので ★開いた 時に 計算し直します★
         ⇒★これで 足りるかは 経営者1 が 実Excel で 数えます★
       ★★測って いない 事★★
         ・★実Excel が この 包みを どう 開くかは ここでは 測れません★
         ・★板の 中の 飾り（罫線・塗り）は 書きません★＝値と 式だけ
         ・★`.xlsb` は 別物★（`workbook.bin` が 2進） */
  function 空き番号(zip, 頭, 尾) {
    var 大 = 0;
    zip.names().forEach(function (n) {
      if (n.indexOf(頭) !== 0) return;
      var 尻 = n.slice(頭.length);
      if (尻.slice(-尾.length) !== 尾) return;
      var v = parseInt(尻.slice(0, 尻.length - 尾.length), 10);
      if (isFinite(v) && v > 大) 大 = v;
    });
    return 大 + 1;
  }
  /** ★`Id="rIdN"` の 一番 大きい N ＋ 1★ */
  function 空きrId(relsXml) {
    var 大 = 0, i = 0, s = String(relsXml || ''), 鍵 = ' Id="rId';
    while (true) {
      var a = s.indexOf(鍵, i);
      if (a < 0) break;
      var b = s.indexOf('"', a + 鍵.length);
      if (b < 0) break;
      var v = parseInt(s.slice(a + 鍵.length, b), 10);
      if (isFinite(v) && v > 大) 大 = v;
      i = b + 1;
    }
    return 'rId' + (大 + 1);
  }
  /** ★`sheetId="N"` の 一番 大きい N ＋ 1★ */
  function 空きsheetId(wbXml) {
    var 大 = 0, i = 0, s = String(wbXml || ''), 鍵 = ' sheetId="';
    while (true) {
      var a = s.indexOf(鍵, i);
      if (a < 0) break;
      var b = s.indexOf('"', a + 鍵.length);
      if (b < 0) break;
      var v = parseInt(s.slice(a + 鍵.length, b), 10);
      if (isFinite(v) && v > 大) 大 = v;
      i = b + 1;
    }
    return 大 + 1;
  }

  /** ★板の xml を 作る★（★値と 式だけ★＝飾りは 書きません）
   *  ★字は `inlineStr` で 書きます★＝★共有文字列を 触らない★
   *    （触ると 元の 板が 指して いる 番号が ずれる 恐れが 在ります） */
  function 板のxml(cells) {
    var 行 = {};
    Object.keys(cells || {}).forEach(function (a) {
      var m = /^([A-Z]+)(\d+)$/.exec(a);
      if (!m) return;
      var r = parseInt(m[2], 10);
      (行[r] = 行[r] || []).push({ 番地: a, 中: cells[a] });
    });
    var 番号 = Object.keys(行).map(Number).sort(function (a, b) { return a - b; });
    var 出 = '';
    番号.forEach(function (r) {
      出 += '<row r="' + r + '">';
      行[r].forEach(function (c) {
        var v = (c.中 && c.中.v !== undefined) ? c.中.v : null;
        var t = c.中 && c.中.t;
        if (v === null || v === '') { 出 += '<c r="' + c.番地 + '"/>'; return; }
        if (t === 's' || typeof v === 'string') {
          出 += '<c r="' + c.番地 + '" t="inlineStr"><is><t xml:space="preserve">'
            + esc(String(v)) + '</t></is></c>';
        } else {
          出 += '<c r="' + c.番地 + '"><v>' + v + '</v></c>';
        }
      });
      出 += '</row>';
    });
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
      + ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
      + '<sheetData>' + 出 + '</sheetData></worksheet>';
  }

  /* ★★名前を `addSheet` に しない★★（2026-09-21）
       `tests/unused-param.test.mjs` は ★関数の 名前で 呼び出しを 数えます★。
       `addSheet` は もう ★2個 在りました★
         `book.html:17339 addSheet()`（画面の ボタン・引数 0）
         `exally-formula.js:34 addSheetToEngine(name)`
       ⇒私の `addSheet(book, name, cells)` を 足したら
         ★「3つ 受けるのに 誰も 1つしか 渡して いない」と 誤報★されました。
       ⇒`板を足す` に しました（★repo に 0個★ と 数えて から 付けました）。
       ★★今日 3回目です★★（`数える` / `読む` / `addSheet`）
       ⇒★名前を 付ける 前に repo に 何個 在るか 数える★ */
  /** ★★板を 1枚 足す★★
   *  @param book  `open()` が 返した 物
   *  @param name  板の 名前
   *  @param cells `{ 'A1': { v, t } }`（`setValues` と 同じ 形・★省いても よい★）
   *  戻り ... `{ 部品, rId, sheetId }`（★足した 物を 名指しで 返します★）
   *  ★もう 在る 名前なら 投げます★（★黙って 2枚に しない★） */
  function 板を足す(book, name, cells) {
    if (!name) throw new Error('板の名前がありません');
    for (var i = 0; i < book.sheets.length; i++) {
      if (book.sheets[i].name === name) throw new Error('もう在る板です: ' + name);
    }
    var zip = book.zip;
    var 番 = 空き番号(zip, 'xl/worksheets/sheet', '.xml');
    var 部品 = 'xl/worksheets/sheet' + 番 + '.xml';
    var rId = 空きrId(book.relsXml);
    var sheetId = 空きsheetId(book.workbookXml);

    /* ①板の 中身 */
    zip.addText(部品, 板のxml(cells));

    /* ②`<sheets>` に 1行 */
    if (book.workbookXml.indexOf('</sheets>') < 0) {
      throw new Error('workbook.xml に <sheets> がありません');
    }
    book.workbookXml = book.workbookXml.replace('</sheets>',
      '<sheet name="' + esc(name) + '" sheetId="' + sheetId + '" r:id="' + rId + '"/></sheets>');

    /* ③rels に 1行 */
    book.relsXml = book.relsXml.replace('</Relationships>',
      '<Relationship Id="' + rId + '"'
      + ' Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"'
      + ' Target="worksheets/sheet' + 番 + '.xml"/></Relationships>');
    zip.replaceText('xl/_rels/workbook.xml.rels', book.relsXml);

    /* ④`[Content_Types].xml` は `save()` で まとめて 直します */
    book.addedTypes = book.addedTypes || [];
    book.addedTypes.push('<Override PartName="/' + 部品 + '"'
      + ' ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');

    book.sheets.push({ name: name, rid: rId, part: 部品 });
    return { 部品: 部品, rId: rId, sheetId: sheetId };
  }

  /** 保存する。★触っていない部品は1バイトも変えずに写る★
   *  戻り: Promise<{ bytes, log }>（log = 書き換えた部品を圧縮できたか） */
  function save(book) {
    var zip = book.zip;
    Object.keys(book.dirty).forEach(function (part) { zip.replaceText(part, book.dirty[part]); });

    // 共有文字列に足した物を書き戻す（件数も増やす）
    if (book.sst && book.sst.xml && book.sst.added.length) {
      var add = book.sst.added.map(function (s) { return '<si><t xml:space="preserve">' + esc(s) + '</t></si>'; }).join('');
      var xml = book.sst.xml.replace(/<\/sst>\s*$/, add + '</sst>');
      var total = book.sst.count + book.sst.added.length;
      xml = xml.replace(/(<sst\b[^>]*?)\scount="\d+"/, '$1 count="' + total + '"')
        .replace(/(<sst\b[^>]*?)\suniqueCount="\d+"/, '$1 uniqueCount="' + total + '"');
      zip.replaceText('xl/sharedStrings.xml', xml);
    }

    // ★開いた時に必ず計算し直させる（保険。本命は <v> を自分で埋めること）★
    var wbx = book.workbookXml;
    if (/<calcPr\b[^>]*\/>/.test(wbx)) {
      wbx = wbx.replace(/<calcPr\b([^>]*?)\s*\/>/, function (_, a) {
        return '<calcPr' + a.replace(/\s+fullCalcOnLoad="[^"]*"/g, '') + ' fullCalcOnLoad="1"/>';
      });
    } else if (/<calcPr\b/.test(wbx)) {
      wbx = wbx.replace(/<calcPr\b([^>]*?)>/, '<calcPr$1 fullCalcOnLoad="1">');
    } else {
      wbx = wbx.replace(/<\/workbook>/, '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>');
    }
    zip.replaceText('xl/workbook.xml', wbx);

    /* ★★足した 板の 型を [Content_Types].xml に 入れます★★（2026-09-21）
         ★ここが 抜けると 実Excel は 修復を 言います★（板は 在るのに 型が 無い） */
    if (book.addedTypes && book.addedTypes.length) {
      return zip.text('[Content_Types].xml').then(function (ct) {
        if (ct.indexOf('</Types>') < 0) throw new Error('[Content_Types].xml が読めません');
        zip.replaceText('[Content_Types].xml',
          ct.replace('</Types>', book.addedTypes.join('') + '</Types>'));
        return zip.build();
      });
    }
    return zip.build();
  }

  /** 中身から見て、どの形式か（拡張子を信じない） */
  function kindOf(zip) {
    if (zip.has('xl/workbook.bin')) return 'xlsb';
    if (zip.has('xl/vbaProject.bin')) return 'xlsm';
    if (zip.has('xl/workbook.xml')) return 'xlsx';
    return 'unknown';
  }

  return { open: open, setValues: setValues, save: save, sheetOf: sheetOf, kindOf: kindOf,
    板を足す: 板を足す, 板のxml: 板のxml };
});
