/* cond-format.js — ★条件付き書式の「当たり判定」だけ（純関数）★
 *
 *  真値 = tests/fixtures/cond-format-golden.json（実Excel 16.0.20228 を COM で動かして測った）
 *  ★測って分かった事（思い込みと違った所）★
 *    ・★文字は どんな数より大きい★  「あいう」は「25より大きい」に ★当たる★
 *    ・★空セルは 0★               「より大きい」に当たらず「より小さい」に ★当たる★
 *    ・★空文字を返す式★ =IF(1=1,"","x") は ★文字扱い★（空セルと違う）
 *    ・★エラー(#DIV/0! 等)は どのルールにも当たらない★
 *    ・★上位/下位は TopBottom 0＝下位 / 1＝上位★
 *    ・★2本 重なったら 番号の小さいルールが勝つ★
 *    ・★手で塗った色は 消えない（条件を外すと戻る）★＝ここでは「上に描く物」を返すだけ
 *    ・★計算には触らない★
 *
 *  ★ここは見た目だけ。値も式も 1バイトも変えない★
 */
(function (root) {
  'use strict';

  var ERR = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!|GETTING_DATA)$/;

  function isError(t) { return ERR.test(String(t).trim()); }
  /** ★うちの計算は =1/0 で Infinity を返す事がある（2026-08-21 実測）＝エラー扱いにする★ */
  function isNotComputed(t) { return /^(-?Infinity|NaN)$/.test(String(t).trim()); }
  function isForcedText(cell) {
    return !!(cell && typeof cell.f === 'string' && cell.f.charAt(0) === "'");
  }

  /**
   * ★セルの「中身」を Excelの比べ方に合わせて1つに畳む★
   * @returns {{型:'数'|'文字'|'空'|'エラー', 数:number|null, 文字:string}}
   */
  function 読む(cell) {
    if (!cell) return { 型: '空', 数: 0, 文字: '' };
    var 生 = cell.v;
    /* ★実物のグリッドは f に「打った字そのもの」を入れる★（10 と打つと {v:'10', f:'10', d:'10'}）。
       ★f が在る＝式 と読むと、ただの数まで 文字扱いになり「25より大きい」に全部 当たる★
       （2026-08-22 本物の画面で押して見つけた。部品だけの検査では 出なかった） */
    var 式 = (typeof cell.f === 'string' && cell.f.charAt(0) === '=');

    if (式) {
      var d = cell.d;
      if (typeof d === 'number' && isFinite(d)) return { 型: '数', 数: d, 文字: String(d) };
      var ds = (d === undefined || d === null) ? '' : String(d);
      if (isError(ds) || isNotComputed(ds)) return { 型: 'エラー', 数: null, 文字: ds };
      /* ★空文字を返す式は 文字扱い（実測③・空セルと違う）★ */
      return { 型: '文字', 数: null, 文字: ds };
    }
    if (isForcedText(cell)) {
      var t0 = String(cell.f).slice(1);
      return { 型: '文字', 数: null, 文字: t0 };
    }
    if (typeof 生 === 'number' && isFinite(生)) return { 型: '数', 数: 生, 文字: String(生) };
    if (生 === undefined || 生 === null || 生 === '') return { 型: '空', 数: 0, 文字: '' };
    var s = String(生);
    if (isError(s) || isNotComputed(s)) return { 型: 'エラー', 数: null, 文字: s };
    var n = Number(s);
    if (s.trim() !== '' && !isNaN(n) && isFinite(n)) return { 型: '数', 数: n, 文字: s };
    return { 型: '文字', 数: null, 文字: s };
  }

  /**
   * ★Excelの並び★ 数 < 文字。空は 0（数）。
   * @returns {number|null} 中身 と しきい値 を比べた -1/0/1（比べられない時は null）
   */
  function 比べる(中身, しきい) {
    if (中身.型 === 'エラー') return null;               /* ★エラーは どれにも当たらない★ */
    var 相手 = (typeof しきい === 'number') ? { 型: '数', 数: しきい, 文字: String(しきい) }
      : 読む({ v: しきい });
    /* ★実測③＝「"" に等しい」は 空セルにも 空文字を返す式にも 当たる★
       （空セルを 0 として比べると「文字より小さい」になって 当たらなくなる） */
    if (typeof しきい === 'string' && しきい === '') {
      if (中身.型 === '空') return 0;
      if (中身.型 === '文字') return 中身.文字 === '' ? 0 : 1;
      return -1;                                       /* 数 は 文字より小さい */
    }
    var 左が文字 = (中身.型 === '文字');
    var 右が文字 = (相手.型 === '文字');
    if (左が文字 && 右が文字) {
      var a = 中身.文字.toLowerCase(), b = 相手.文字.toLowerCase();
      return a < b ? -1 : (a > b ? 1 : 0);
    }
    if (左が文字 && !右が文字) return 1;                 /* ★文字は どんな数より大きい（実測①）★ */
    if (!左が文字 && 右が文字) return -1;
    var x = 中身.数, y = 相手.数;                        /* 空は 0（実測②） */
    return x < y ? -1 : (x > y ? 1 : 0);
  }

  /**
   * ★1つのセルが そのルールに当たるか★
   * @param {Object} cell   セル（{v,f,d}）
   * @param {Object} ルール { 種類, 演算, 値1, 値2 }
   * @param {Object} 文脈   { 範囲の値たち:[…], 式で判定:function(cell)->bool }
   */
  function 当たるか(cell, ルール, 文脈) {
    ルール = ルール || {};
    文脈 = 文脈 || {};
    var 中身 = 読む(cell);
    if (中身.型 === 'エラー') return false;              /* ★エラーは どのルールにも当たらない★ */

    switch (ルール.種類) {
      case 'セルの値': {
        var c = 比べる(中身, ルール.値1);
        if (c === null) return false;
        switch (ルール.演算) {
          case 'より大きい': return c > 0;
          case 'より小さい': return c < 0;
          case '以上':       return c >= 0;
          case '以下':       return c <= 0;
          case '等しい':     return c === 0;
          case '等しくない': return c !== 0;
          case '範囲内': {
            var c2 = 比べる(中身, ルール.値2);
            if (c2 === null) return false;
            return c >= 0 && c2 <= 0;
          }
          default: return false;
        }
      }
      case '文字を含む': {
        if (中身.型 === '空') return false;
        var 探す = String(ルール.値1 === undefined || ルール.値1 === null ? '' : ルール.値1);
        if (探す === '') return false;
        /* ★実測⑥＝大文字小文字を区別しない・途中に在ってもよい★ */
        return 中身.文字.toLowerCase().indexOf(探す.toLowerCase()) >= 0;
      }
      case '重複する値':
      case '一意の値': {
        if (中身.型 === '空') return false;
        var 数え = 0;
        var 並び = 文脈.範囲の値たち || [];
        for (var i = 0; i < 並び.length; i++) {
          var o = 読む(並び[i]);
          if (o.型 === '空' || o.型 === 'エラー') continue;
          if (同じか(o, 中身)) 数え++;
        }
        return ルール.種類 === '重複する値' ? 数え >= 2 : 数え === 1;
      }
      case '上位下位': {
        if (中身.型 !== '数') return false;              /* 文字・空は 順位に入らない */
        var 数たち = [];
        var 並び2 = 文脈.範囲の値たち || [];
        for (var j = 0; j < 並び2.length; j++) {
          var o2 = 読む(並び2[j]);
          if (o2.型 === '数') 数たち.push(o2.数);
        }
        if (!数たち.length) return false;
        var n = Math.max(1, Number(ルール.値1) || 10);
        数たち.sort(function (a, b) { return a - b; });
        if (ルール.演算 === '下位') {
          var 境下 = 数たち[Math.min(n, 数たち.length) - 1];
          return 中身.数 <= 境下;
        }
        var 境上 = 数たち[Math.max(0, 数たち.length - n)];
        return 中身.数 >= 境上;
      }
      case '式': {
        /* ★式のルールは 計算する物を持っていない（ここは純関数）★
           呼ぶ側が 文脈.式で判定 を渡す＝本物の計算エンジンで測る。
           ★渡されなければ 当てない（決め打ちしない）★ */
        if (typeof 文脈.式で判定 !== 'function') return false;
        return !!文脈.式で判定(cell);
      }
      default: return false;
    }
  }

  function 同じか(a, b) {
    if (a.型 !== b.型) {
      /* 数の 10 と 文字の "10" は 別物（Excelの重複も 別物として扱う） */
      return false;
    }
    if (a.型 === '数') return a.数 === b.数;
    return a.文字.toLowerCase() === b.文字.toLowerCase();
  }

  /**
   * ★そのセルに 実際に描く書式★（番号の小さいルールが勝つ）
   * @param {Object} cell
   * @param {Array}  ルールたち  番号順（先頭が一番強い）。各 { …, 書式:{塗り,文字色,太字} }
   * @param {Object} 文脈
   * @returns {Object|null} { 塗り, 文字色, 太字 }（当たらなければ null）
   */
  function 効く書式(cell, ルールたち, 文脈) {
    var 出 = null;
    var 並び = ルールたち || [];
    for (var i = 0; i < 並び.length; i++) {
      if (!当たるか(cell, 並び[i], 文脈)) continue;
      var f = 並び[i].書式 || {};
      出 = 出 || {};
      /* ★番号の小さい方が勝つ＝先に決まった物は 上書きしない★
         ★項目が違えば 混ざる（塗りだけのルールと 文字色だけのルール）★ */
      if (出.塗り === undefined && f.塗り) 出.塗り = f.塗り;
      if (出.文字色 === undefined && f.文字色) 出.文字色 = f.文字色;
      if (出.太字 === undefined && f.太字 !== undefined) 出.太字 = !!f.太字;
    }
    return 出;
  }

  /** ★今 何個 当たるか（押す前に見せる数）★ */
  function 当たる数(ルール, セルたち, 文脈) {
    var n = 0;
    var 並び = セルたち || [];
    var c = 文脈 || { 範囲の値たち: 並び };
    if (!c.範囲の値たち) c.範囲の値たち = 並び;
    for (var i = 0; i < 並び.length; i++) if (当たるか(並び[i], ルール, c)) n++;
    return n;
  }

  /**
   * ★式のルールを そのセルの場所まで ずらす★
   *  実測⑤＝式は「範囲の左上のセル」から見て書く。$ で止めた所は動かない。
   *  例）=$A1>25 を A1:C3 に当てる → 3行目のセルでは =$A3>25 で判定される
   *  ★引用符の中（"A1" のような文字）は ずらさない★
   */
  function 式をずらす(式, 行ずれ, 列ずれ) {
    var src = String(式 === undefined || 式 === null ? '' : 式);
    if (!行ずれ && !列ずれ) return src;
    var 出 = '', 中 = false, i = 0;
    while (i < src.length) {
      var ch = src.charAt(i);
      if (ch === '"') { 中 = !中; 出 += ch; i++; continue; }
      if (中) { 出 += ch; i++; continue; }
      var m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})/.exec(src.slice(i));
      if (m && !/[A-Za-z0-9_$.]/.test(src.charAt(i - 1) || '')) {
        var 列 = m[1] ? m[2] : 列文字(列番号(m[2]) + 列ずれ);
        var 行 = m[3] ? m[4] : String(Number(m[4]) + 行ずれ);
        if (列 === null || Number(行) < 1) { 出 += '#REF!'; }
        else { 出 += m[1] + 列 + m[3] + 行; }
        i += m[0].length;
        continue;
      }
      出 += ch; i++;
    }
    return 出;
  }

  function 列番号(s) {
    var n = 0;
    var t = String(s).toUpperCase();
    for (var i = 0; i < t.length; i++) n = n * 26 + (t.charCodeAt(i) - 64);
    return n - 1;                                    /* A=0 */
  }
  function 列文字(n) {
    if (n < 0) return null;
    var s = '';
    n = n + 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /** Excelの既定の見た目（実測値・「濃い赤の文字、明るい赤の背景」） */
  var 既定の書式 = { 塗り: '#FFC7CE', 文字色: '#9C0006', 太字: false };

  var api = {
    読む: 読む, 比べる: 比べる, 当たるか: 当たるか, 効く書式: 効く書式, 式をずらす: 式をずらす,
    当たる数: 当たる数, 既定の書式: 既定の書式,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.CondFormat = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
