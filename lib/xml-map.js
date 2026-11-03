/* xml-map.js — ★XML の 対応付け・読み込み・書き出し★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-xml.ps1 / tools/measure-xml2.ps1
 *
 *  ● 対応付け（XmlMaps）
 *      `XmlMaps.Add(ファイル, 'people')` で 作れる
 *      名前 … ★`people_対応付け`★（★ルートの 名前 ＋ `_対応付け`★）
 *      `RootElementName` … `people`／`Schemas.Count` … 1
 *      ★`IsExportable` は 対応付ける前 `False`／セルに 結んだ後 `True`★
 *
 *  ● 読み込み（Import）
 *      ★セルに 結ぶ前に `Import` を 呼ぶと 断られる★
 *        「対応付けられている要素がないためデータはインポートされませんでした。
 *          Range.XPath.SetValue を使用して XML 要素をシートに対応付けてください。」
 *      `Range('A1').XPath.SetValue(map, '/people/person/name')` で 結ぶ
 *        → `Range('A1').XPath.Value` が `/people/person/name` に なる
 *      結んでから `Import` → ★戻りは 0（成功）★
 *      ★入った 中身は A1='ta' B1='20' の ★1行だけ★★
 *        （くり返す 物を 何行も 入れるには 実Excel は ★表（ListObject）が 要る★）
 *
 *  ● 書き出し（Export）
 *      `IsExportable` が False だと ★断られる★
 *        「対応付けがエクスポートできなかったため…IsExportable を使用して確認してください。」
 *      出た XML … ★`standalone="yes"` が 付く★／
 *        ルートに ★`xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"` が 足される★
 *      結んでいない セル（A2）を 直しても ★出る 中身は 変わらない★
 *
 *  ● 対応付けの 決め（実測の 既定）
 *      AdjustColumnWidth=True ／ ★AppendOnImport=False★ ／ PreserveColumnFilter=True ／
 *      PreserveNumberFormatting=True ／ SaveDataSourceDefinition=True ／
 *      ★ShowImportExportValidationErrors=False★
 *
 *  ★違う所（わざと）★
 *    実Excel は くり返す 物を 何行も 入れるのに ★表（ListObject）が 要る★。
 *    ★うちは 結んだ XPath が くり返す 物なら そのまま 何行も 入れる★
 *    （表を 先に 作らせない＝人の 手間を 減らす）。★実測の 1行だけ★との 違いは ここに 書く。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.XmlMap = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★実測の 既定★（そのまま 同じに する） */
  var 既定の決め = {
    列の幅を合わせる: true,          /* AdjustColumnWidth = True */
    読み込みで足す: false,           /* ★AppendOnImport = False★ */
    絞り込みを残す: true,            /* PreserveColumnFilter = True */
    書式を残す: true,                /* PreserveNumberFormatting = True */
    元の場所を覚える: true,          /* SaveDataSourceDefinition = True */
    検査の間違いを見せる: false,     /* ★ShowImportExportValidationErrors = False★ */
  };

  /** 対応付けの 名前（★実測＝ルート ＋ `_対応付け`★） */
  function 名前を決める(ルート名) {
    return String(ルート名 || 'xml') + '_対応付け';
  }

  /* ── XML を 読む（小さな 読み取り。ブラウザの DOMParser が 在れば そちら）── */
  function 木にする(字) {
    if (typeof DOMParser !== 'undefined') {
      var d = new DOMParser().parseFromString(String(字), 'application/xml');
      var err = d.getElementsByTagName('parsererror');
      if (err && err.length) return null;
      return _DOMを木に(d.documentElement);
    }
    return _手で木に(String(字));
  }
  function _DOMを木に(節) {
    if (!節) return null;
    var 出 = { 名: 節.nodeName, 字: '', 子: [] };
    for (var i = 0; i < 節.childNodes.length; i++) {
      var c = 節.childNodes[i];
      if (c.nodeType === 1) 出.子.push(_DOMを木に(c));
      else if (c.nodeType === 3) 出.字 += c.nodeValue;
    }
    出.字 = 出.字.trim();
    return 出;
  }
  /** ★ブラウザの 外（試験）でも 動く 小さな 読み取り★
   *  ★ブラウザの DOMParser と 同じ 所で 断る★＝
   *    閉じ札の 名前違い・閉じ忘れ・根が 2つ は null（08-30 実測で 食い違っていた）。 */
  function _手で木に(s) {
    s = s.replace(/<\?[\s\S]*?\?>/g, '').replace(/<!--[\s\S]*?-->/g, '');
    var 位 = 0;
    var 積 = [];
    var 根 = null;
    var re = /<\/?([A-Za-z_][\w.:-]*)([^>]*?)(\/?)>/g;
    var m;
    while ((m = re.exec(s)) !== null) {
      var 字 = s.slice(位, m.index).trim();
      if (字 && 積.length) 積[積.length - 1].字 += 字;
      位 = re.lastIndex;
      var 閉じ = m[0].charAt(1) === '/';
      var 自己 = m[3] === '/';
      if (閉じ) {
        /* ★閉じ札は 開いた 札と 同じ 名前で ないと 駄目★
           （ブラウザの 読み取りは 断る。★2つの 道で 答えが 違うと 嘘に なる★ので 合わせる） */
        if (!積.length || 積[積.length - 1].名 !== m[1]) return null;
        積.pop();
        continue;
      }
      if (根 && !積.length) return null;      /* 根が 2つ＝XML では ない */
      var 節 = { 名: m[1], 字: '', 子: [] };
      if (積.length) 積[積.length - 1].子.push(節); else 根 = 節;
      if (!自己) 積.push(節);
    }
    if (積.length) return null;               /* ★閉じ忘れ★ */
    return 根;
  }

  /** その 木に 在る 道（XPath）を 全部 出す（★くり返す 物は 1つに まとめる★） */
  function 道を集める(木) {
    var 出 = [];
    var 見た = {};
    (function 歩く(節, 道) {
      if (!節) return;
      var 今 = 道 + '/' + 節.名;
      var 葉 = !節.子.length;
      if (!見た[今]) {
        見た[今] = true;
        出.push({ 道: 今, 葉: 葉, くり返す: false });
      }
      /* 同じ 名前の 子が 2つ 以上＝くり返す */
      var 数 = {};
      for (var i = 0; i < 節.子.length; i++) 数[節.子[i].名] = (数[節.子[i].名] || 0) + 1;
      for (var j = 0; j < 節.子.length; j++) {
        歩く(節.子[j], 今);
        if (数[節.子[j].名] > 1) {
          for (var k = 0; k < 出.length; k++) {
            if (出[k].道.indexOf(今 + '/' + 節.子[j].名) === 0) 出[k].くり返す = true;
          }
        }
      }
    }(木, ''));
    return 出;
  }

  /** その 道に 当てはまる 値を 全部 取る（★くり返す 物は 何個も 返る★） */
  function 値を取る(木, 道) {
    var 部 = String(道 || '').split('/').filter(function (v) { return v; });
    if (!木 || !部.length) return [];
    if (木.名 !== 部[0]) return [];
    var いま = [木];
    for (var i = 1; i < 部.length; i++) {
      var 次 = [];
      for (var j = 0; j < いま.length; j++) {
        for (var k = 0; k < いま[j].子.length; k++) {
          if (いま[j].子[k].名 === 部[i]) 次.push(いま[j].子[k]);
        }
      }
      いま = 次;
    }
    return いま.map(function (v) { return v.字; });
  }

  /** ★読み込み★＝結んだ 道ごとに 値を 並べる
   *  @param 結び [{道, r, c}]（セルの 場所）
   *  @returns [{r, c, 値}]（★くり返す 物は 下に 伸びる★）
   */
  function 読み込む(木, 結び) {
    var 出 = [];
    for (var i = 0; i < (結び || []).length; i++) {
      var b = 結び[i];
      var 値たち = 値を取る(木, b.道);
      for (var j = 0; j < 値たち.length; j++) {
        出.push({ r: b.r + j, c: b.c, 値: 値たち[j] });
      }
    }
    return 出;
  }

  /** ★出せるか★（実測＝結ぶ前 False・結んだ後 True） */
  function 出せるか(結び) {
    return !!(結び && 結び.length);
  }

  /** ★書き出し★＝結んだ セルの 値から XML を 組み立て直す
   *  実測に 合わせて ★standalone="yes"★ と ★xmlns:xsi★ を 付ける。
   *  @param 読む (r,c) → 値
   */
  function 書き出す(木, 結び, 読む) {
    if (!出せるか(結び)) return null;                 /* ★実Excel も 断る★ */
    var ルート = 木 ? 木.名 : 'root';
    /* 何行 分 在るか（一番 長い 道に 合わせる） */
    var 行数 = 0;
    for (var i = 0; i < 結び.length; i++) {
      var n = 0;
      while (n < 10000) {
        var v = 読む(結び[i].r + n, 結び[i].c);
        if (v === undefined || v === null || v === '') break;
        n++;
      }
      if (n > 行数) 行数 = n;
    }
    if (!行数) 行数 = 1;
    /* 結びを 「くり返す 親」ごとに まとめる */
    var 親 = null;
    for (var j = 0; j < 結び.length; j++) {
      var 部 = 結び[j].道.split('/').filter(function (v) { return v; });
      var p = '/' + 部.slice(0, 部.length - 1).join('/');
      if (親 === null) 親 = p;
    }
    var 中 = 親 ? 親.split('/').filter(function (v) { return v; }) : [];
    var 子名 = 中.length > 1 ? 中[中.length - 1] : null;

    var 行 = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<' + ルート + ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'];
    for (var r = 0; r < 行数; r++) {
      if (子名) 行.push('\t<' + 子名 + '>');
      for (var k = 0; k < 結び.length; k++) {
        var 名 = 結び[k].道.split('/').pop();
        var 値 = 読む(結び[k].r + r, 結び[k].c);
        if (値 === undefined || 値 === null) 値 = '';
        行.push((子名 ? '\t\t' : '\t') + '<' + 名 + '>' + _安全(値) + '</' + 名 + '>');
      }
      if (子名) 行.push('\t</' + 子名 + '>');
    }
    行.push('</' + ルート + '>');
    return 行.join('\n');
  }
  function _安全(v) {
    return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  return {
    既定の決め: 既定の決め, 名前を決める: 名前を決める,
    木にする: 木にする, 道を集める: 道を集める, 値を取る: 値を取る,
    読み込む: 読み込む, 出せるか: 出せるか, 書き出す: 書き出す,
  };
}));
