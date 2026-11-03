/* named-ranges.js — ★名前の定義（名前付き範囲）★ 2026-08-29
 *
 *  ★なぜ 自分で 作るか（実測 2026-08-29）★
 *    HyperFormula にも 名前の仕組みは 在るが ★日本語の名前を 受け付けない★:
 *        uriage      … ok
 *        売上 / うりあげ / _売上 / 売上_2026 / Sales1 … ★全部 「invalid」★
 *    司さんの実物は ★日本語の名前★（R8.1 / 歩合 …）で 書かれている。
 *    ⇒ ★エンジンに 渡す前に「名前 → 実際の範囲」に 開く★。
 *      開く所は ★convertFormula の中 1か所★（エンジンへ入る 唯一の門）。
 *
 *  ★実Excelの 決まり（COMで 実測 2026-08-29）★
 *    ・=SUM(うりあげ) → 60（日本語の名前は ★使える★）
 *    ・★A1 のような 番地の形は 名前に できない★（Excelが 断る）
 *    ・★数で 始まる名前も できない★（1あ → 断られた）
 *    ・ブック全体の名前と ★シートだけの名前★（Sheet1!しーとない）が ある
 *
 *  ★開く時に 気をつける事★
 *    ・★字の中（"…"）は 触らない★＝ "売上" という 文字は そのまま
 *    ・★長い名前から 先に 開く★（売上 と 売上合計 が 両方 在る時、短い方から だと 壊れる）
 *    ・★他の名前の 一部を 食わない★＝前後が 名前に使える字なら 開かない
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.NamedRanges = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* 名前に 使える字＝英数字・下線・日本語（Excelと 同じ範囲に 寄せる） */
  var 名前の字 = /[A-Za-z0-9_぀-ヿ㐀-鿿ｦ-ﾟ.]/;

  /** ★この名前は 使えるか★（実Excelの 決まりに 合わせる）
   *  返り値: null＝使える ／ 文字列＝断る理由 */
  function 名前を確かめる(名) {
    var s = String(名 == null ? '' : 名).trim();
    if (!s) return '名前が 空です';
    if (s.length > 255) return '名前が 長すぎます（255文字まで）';
    if (/^[0-9]/.test(s)) return '数で 始まる名前は 使えません（実Excelも 断ります）';
    if (/^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$/.test(s)) return 'A1 のような 番地の形は 使えません（実Excelも 断ります）';
    if (/^[RrCc][0-9]*$/.test(s)) return 'R や C だけの名前は 使えません';
    for (var i = 0; i < s.length; i++) {
      if (!名前の字.test(s.charAt(i))) return '使えない字が 入っています: ' + s.charAt(i);
    }
    return null;
  }

  /** 一覧を 持つ（ブック1つぶん） */
  function 作る() {
    var 一覧 = [];   /* { 名, 参照先, シート(null=ブック全体) } */

    function 足す(名, 参照先, シート) {
      var なぜ = 名前を確かめる(名);
      if (なぜ) return { ok: false, なぜ: なぜ };
      if (!String(参照先 || '').trim()) return { ok: false, なぜ: '参照先が 空です' };
      var 同じ = 探す(名, シート);
      if (同じ) { 同じ.参照先 = String(参照先); return { ok: true, 直した: true }; }
      一覧.push({ 名: String(名).trim(), 参照先: String(参照先), シート: シート || null });
      return { ok: true, 足した: true };
    }
    function 探す(名, シート) {
      var s = String(名).trim();
      /* ★シートだけの名前が 先★（Excelと 同じ＝近い方が 勝つ） */
      for (var i = 0; i < 一覧.length; i++) {
        if (一覧[i].名 === s && 一覧[i].シート && 一覧[i].シート === シート) return 一覧[i];
      }
      for (var j = 0; j < 一覧.length; j++) {
        if (一覧[j].名 === s && !一覧[j].シート) return 一覧[j];
      }
      return null;
    }
    function 消す(名, シート) {
      for (var i = 0; i < 一覧.length; i++) {
        if (一覧[i].名 === String(名).trim() && (一覧[i].シート || null) === (シート || null)) {
          一覧.splice(i, 1); return true;
        }
      }
      return false;
    }
    function 全部() { return 一覧.slice(); }
    function 空にする() { 一覧.length = 0; }

    /** ★式の中の 名前を 実際の範囲に 開く★
     *  ・字の中（"…"）は 触らない
     *  ・長い名前から 先に 開く
     *  ・前後が 名前に使える字なら 開かない（他の名前の 一部を 食わない） */
    function 開く(式, シート) {
      var f = String(式 == null ? '' : 式);
      if (!f || !一覧.length) return f;
      /* このシートで 見える名前だけ（長い順） */
      var 見える = [];
      for (var i = 0; i < 一覧.length; i++) {
        var n = 一覧[i];
        if (!n.シート || n.シート === シート) 見える.push(n);
      }
      見える.sort(function (a, b) { return b.名.length - a.名.length; });
      if (!見える.length) return f;

      /* 字の中を 避けながら 1文字ずつ 進む */
      var out = '', i2 = 0, n2 = f.length;
      while (i2 < n2) {
        var c = f.charAt(i2);
        if (c === '"' || c === "'") {
          var 端 = i2 + 1;
          while (端 < n2) {
            if (f.charAt(端) === c) { if (f.charAt(端 + 1) === c) { 端 += 2; continue; } 端++; break; }
            端++;
          }
          out += f.slice(i2, 端); i2 = 端; continue;
        }
        var 当たり = null;
        for (var k = 0; k < 見える.length; k++) {
          var 名 = 見える[k].名;
          if (f.substr(i2, 名.length) !== 名) continue;
          var 前 = i2 > 0 ? f.charAt(i2 - 1) : '';
          var 後 = f.charAt(i2 + 名.length);
          if (前 && 名前の字.test(前)) continue;        // 他の名前の 途中
          if (後 && 名前の字.test(後)) continue;
          if (後 === '(') continue;                      // 関数の名前（=SUM( …）は 触らない
          if (前 === '!') continue;                      // シート名の 直後は 触らない
          当たり = 見える[k]; break;
        }
        if (当たり) {
          out += '(' + String(当たり.参照先).replace(/^=/, '') + ')';
          i2 += 当たり.名.length;
          continue;
        }
        out += c; i2++;
      }
      return out;
    }

    return { 足す: 足す, 探す: 探す, 消す: 消す, 全部: 全部, 空にする: 空にする, 開く: 開く };
  }

  return { 作る: 作る, 名前を確かめる: 名前を確かめる, _名前の字: 名前の字 };
}));
