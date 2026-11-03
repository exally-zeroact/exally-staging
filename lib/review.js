/* review.js — ★校閲（文章校正・アクセシビリティ・変更内容）★ 2026-08-30
 *
 *  ★真値（実Excel 16.0 で 実測）★ … tools/measure-review2.ps1
 *
 *  ● スペル チェック（`Application.CheckSpelling`）
 *      `recieve` → ★False★／`receive` → True
 *      `teh` → ★False★／`the` → True
 *      `apple` → True／`こんにちは` → ★True★（★日本語は いつも 通る★）
 *      ★直しの 案（GetSpellingSuggestions）は COM から 呼べなかった★
 *
 *  ● アクセシビリティ チェック
 *      図形の 代替テキスト（`AlternativeText`）… ★はじめは 空★／入れられる
 *      `Title` も 在る
 *
 *  ● 変更内容を表示 / ブックの共有
 *      `MultiUserEditing` … ★False★（共有していない）
 *      `KeepChangeHistory` … True ／ `ChangeHistoryDuration` … ★0★
 *      `HighlightChangesOnScreen` … False
 *      ★`HighlightChangesOptions` は 共有していないと 通らない（0x800A03EC）★
 *
 *  ● 類義語辞典・翻訳
 *      ★COM から 呼べない★（画面の 窓／Microsoft の サービス）
 *
 *  ★違う所（わざと・出来ない事は 出来ないと 書く）★
 *    ・うちは ★大きな 辞書を 持たない★。
 *      スペルは ★よく まちがえる 英語の 一覧★ ＋
 *      ★同じ ブックの 中で 1文字違いの 言葉★（打ちまちがい）を 見つける。
 *      ★日本語は 実Excel と 同じで いつも 通す★。
 *    ・類義語・翻訳は ★辞書と 外の サービスが 要る★ので うちの AI に つなぐ。
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.Review = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ★よく まちがえる 英語★（実測で 確かめた 2つを 含む） */
  var よくある間違い = {
    recieve: 'receive', teh: 'the', adress: 'address', occured: 'occurred',
    seperate: 'separate', definately: 'definitely', accomodate: 'accommodate',
    neccessary: 'necessary', occurence: 'occurrence', publically: 'publicly',
    recomend: 'recommend', refered: 'referred', succesful: 'successful',
    untill: 'until', wich: 'which', beleive: 'believe', calender: 'calendar',
    concious: 'conscious', embarass: 'embarrass', existance: 'existence',
    goverment: 'government', independant: 'independent', maintainance: 'maintenance',
    occassion: 'occasion', persistant: 'persistent', priviledge: 'privilege',
    recieved: 'received', suprise: 'surprise', tommorow: 'tomorrow',
    truely: 'truly', wierd: 'weird', writting: 'writing',
  };

  var 日本語 = /[぀-ヿ一-鿿]/;

  /** 字を 言葉に 分ける（英語の 語だけ 拾う） */
  function 言葉に分ける(字) {
    var s = String(字 === null || 字 === undefined ? '' : 字);
    var 出 = s.match(/[A-Za-z][A-Za-z']*/g);
    return 出 || [];
  }

  /** 1文字ずつの 違いを 数える（打ちまちがい 探しに 使う） */
  function 違いの数(a, b) {
    a = String(a).toLowerCase(); b = String(b).toLowerCase();
    if (a === b) return 0;
    if (Math.abs(a.length - b.length) > 1) return 2;   /* 2以上は もう 見ない */
    var 前 = [];
    for (var j = 0; j <= b.length; j++) 前.push(j);
    for (var i = 1; i <= a.length; i++) {
      var 今 = [i];
      for (var k = 1; k <= b.length; k++) {
        今.push(Math.min(前[k] + 1, 今[k - 1] + 1,
          前[k - 1] + (a.charAt(i - 1) === b.charAt(k - 1) ? 0 : 1)));
      }
      前 = 今;
    }
    return 前[b.length];
  }

  /** ★スペルを 見る★
   *  @param 字たち セルの 中身の 配列（{値, 場所} でも 字だけでも 良い）
   *  @returns [{語, 場所, 訳, 直し}]
   */
  function スペルを見る(字たち) {
    var 出 = [];
    var 数 = {};
    var 場所 = {};
    for (var i = 0; i < (字たち || []).length; i++) {
      var x = 字たち[i];
      var 値 = (x && typeof x === 'object') ? x.値 : x;
      var 所 = (x && typeof x === 'object') ? x.場所 : '';
      if (日本語.test(String(値 === null || 値 === undefined ? '' : 値))) continue;  /* ★日本語は 通す★ */
      var 語たち = 言葉に分ける(値);
      for (var j = 0; j < 語たち.length; j++) {
        var 語 = 語たち[j];
        var 小 = 語.toLowerCase();
        数[小] = (数[小] || 0) + 1;
        if (!場所[小]) 場所[小] = 所;
        if (Object.prototype.hasOwnProperty.call(よくある間違い, 小)) {
          出.push({ 語: 語, 場所: 所, 訳: 'よく まちがえる 言葉', 直し: よくある間違い[小] });
        }
      }
    }
    /* ★同じ ブックの 中で 1文字違い★＝打ちまちがいの 目印
       （たくさん 出て くる 方が 正しい と 見る） */
    var 語一覧 = Object.keys(数);
    for (var a = 0; a < 語一覧.length; a++) {
      for (var b = 0; b < 語一覧.length; b++) {
        if (a === b) continue;
        var 少 = 語一覧[a], 多 = 語一覧[b];
        if (少.length < 4) continue;                 /* 短い 語は 見ない（別の 語が 多い） */
        if (数[少] !== 1 || 数[多] < 3) continue;     /* 1回だけ vs 3回以上 */
        if (違いの数(少, 多) !== 1) continue;
        if (Object.prototype.hasOwnProperty.call(よくある間違い, 少)) continue;  /* もう 出した */
        出.push({ 語: 少, 場所: 場所[少], 訳: '同じ ブックに 1文字違いの 言葉が ' + 数[多] + '回',
                  直し: 多 });
      }
    }
    return 出;
  }

  /** ★アクセシビリティを 見る★（実Excel の 検査の 主な 中身と 同じ ところ）
   *  @param o {物たち, 結合たち, シート名たち, 色の組たち}
   *  @returns [{程度:'赤'|'黄', 何, どこ, どうする}]
   */
  function 見やすさを見る(o) {
    o = o || {};
    var 出 = [];
    var 物 = o.物たち || [];
    for (var i = 0; i < 物.length; i++) {
      var v = 物[i];
      if (!v.代替 || !String(v.代替).trim()) {
        出.push({ 程度: '赤', 何: '絵や 図に 説明が 無い',
          どこ: v.名 || ('物' + (i + 1)),
          どうする: '目が 見えない 人にも 分かるよう 何の 絵か 一言 入れてください' });
      }
    }
    var 結合 = o.結合たち || [];
    for (var j = 0; j < 結合.length; j++) {
      出.push({ 程度: '黄', 何: 'セルを 結合している', どこ: 結合[j],
        どうする: '読み上げの 順が 分からなく なります。出来れば 結合を やめてください' });
    }
    var 名 = o.シート名たち || [];
    for (var k = 0; k < 名.length; k++) {
      if (/^Sheet\d+$/.test(名[k]) || /^シート\d+$/.test(名[k])) {
        出.push({ 程度: '黄', 何: 'シートの 名前が そのまま', どこ: 名[k],
          どうする: '何の 表か 分かる 名前に 変えてください' });
      }
    }
    var 組 = o.色の組たち || [];
    for (var m = 0; m < 組.length; m++) {
      var 比 = 明るさの比(組[m].字, 組[m].背);
      if (比 < 4.5) {
        出.push({ 程度: '赤', 何: '字と 背中の 色が 近すぎる（' + 比.toFixed(1) + ':1）',
          どこ: 組[m].場所,
          どうする: '4.5:1 以上に してください（読めない 人が 出ます）' });
      }
    }
    return 出;
  }

  /** 明るさの 比（見やすさの 決まりで 使う 数え方） */
  function 明るさの比(色1, 色2) {
    var a = 明るさ(色1), b = 明るさ(色2);
    var 上 = Math.max(a, b), 下 = Math.min(a, b);
    return (上 + 0.05) / (下 + 0.05);
  }
  function 明るさ(色) {
    var s = String(色 || '#000000').replace('#', '');
    if (s.length === 3) s = s[0] + s[0] + s[1] + s[1] + s[2] + s[2];
    var v = [0, 2, 4].map(function (i) {
      var c = parseInt(s.slice(i, i + 2), 16) / 255;
      return c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
    });
    return 0.2126 * v[0] + 0.7152 * v[1] + 0.0722 * v[2];
  }

  return {
    よくある間違い: よくある間違い, 言葉に分ける: 言葉に分ける, 違いの数: 違いの数,
    スペルを見る: スペルを見る, 見やすさを見る: 見やすさを見る,
    明るさの比: 明るさの比, 明るさ: 明るさ,
  };
}));
