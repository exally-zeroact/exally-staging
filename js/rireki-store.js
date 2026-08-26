/* rireki-store.js — ★履歴の置き場＝倉庫（会社の物）★
 *
 *  ★決まり（司さん 2026-08-25／指示役 08-26）★
 *    ・★倉庫に置く★＝端末を変えても 残る。★競合は 客のブラウザの中(IndexedDB)だけ★＝端末を変えると消える。
 *    ・★客のブックには タブを1つも足さない★（競合は「Claude Log」タブを作る。真似しない）
 *    ・★本人の物だけ 見える★＝RLS（倉庫の側で 締める）
 *    ・★倉庫が使えない時も 画面は動く★（その時だけ この画面の中に持つ＝消える事は 客に言う）
 *
 *  ★向き先は js/supa-config.js だけが持つ★（この画面は 名前を持たない）
 */
(function (root) {
  'use strict';

  var 部屋 = 'exally';
  var 表 = 'rireki';
  var 仮置き = [];              /* ★倉庫が使えない時の 逃げ場（この画面を閉じると消える）★ */
  var 倉庫が使えるか = null;    /* null=まだ分からない（★0件・異常なしにしない★） */

  function つなぎ() {
    var SUPA = root.SUPA;
    if (!SUPA || !SUPA.url || !SUPA.key) return null;
    var sb = (root.Auth && root.Auth.sb) || null;
    if (!sb) return null;
    return { url: String(SUPA.url).replace(/\/+$/, ''), key: SUPA.key, sb: sb };
  }

  function 頭(つ, tk, 足す) {
    var h = {
      apikey: つ.key,
      'Accept-Profile': 部屋,
      'Content-Profile': 部屋,
      'Content-Type': 'application/json',
    };
    if (tk) h.Authorization = 'Bearer ' + tk;
    for (var k in (足す || {})) if (Object.prototype.hasOwnProperty.call(足す, k)) h[k] = 足す[k];
    return h;
  }

  function 合言葉(つ) {
    return つ.sb.auth.getSession().then(function (r) {
      return (r && r.data && r.data.session && r.data.session.access_token) || null;
    }).catch(function () { return null; });
  }

  /** ★1件 残す★（倉庫が駄目でも 画面は止めない） */
  function 残す(件) {
    var 物 = {
      itsu: new Date().toISOString(),
      shurui: String(件.種類 || 'その他'),
      midashi: String(件.見出し || ''),
      nakami: 件.中身 || {},
      file_name: 件.ファイル名 ? String(件.ファイル名) : null,
      credit: Number(件.クレジット || 0),
    };
    var つ = つなぎ();
    if (!つ) { 仮置き.push(物); 倉庫が使えるか = false; return Promise.resolve(false); }
    return 合言葉(つ).then(function (tk) {
      if (!tk) { 仮置き.push(物); 倉庫が使えるか = false; return false; }
      return fetch(つ.url + '/rest/v1/' + 表, {
        method: 'POST',
        headers: 頭(つ, tk, { Prefer: 'return=minimal' }),
        body: JSON.stringify(物),
      }).then(function (res) {
        倉庫が使えるか = !!res.ok;
        if (!res.ok) 仮置き.push(物);
        return !!res.ok;
      });
    }).catch(function () { 仮置き.push(物); 倉庫が使えるか = false; return false; });
  }

  /** ★読む★（倉庫→仮置きの順に混ぜる。★倉庫が読めない事は 隠さない★） */
  function 読む(opt) {
    opt = opt || {};
    var 上限 = opt.上限 || 200;
    var つ = つなぎ();
    var 仮 = 仮置き.map(直す);
    if (!つ) { 倉庫が使えるか = false; return Promise.resolve({ 行たち: 仮, 倉庫: false }); }
    return 合言葉(つ).then(function (tk) {
      if (!tk) { 倉庫が使えるか = false; return { 行たち: 仮, 倉庫: false }; }
      var u = つ.url + '/rest/v1/' + 表 + '?select=*&order=itsu.desc&limit=' + 上限;
      return fetch(u, { headers: 頭(つ, tk) }).then(function (res) {
        if (!res.ok) { 倉庫が使えるか = false; return { 行たち: 仮, 倉庫: false }; }
        return res.json().then(function (a) {
          倉庫が使えるか = true;
          return { 行たち: (a || []).map(直す).concat(仮), 倉庫: true };
        });
      });
    }).catch(function () { 倉庫が使えるか = false; return { 行たち: 仮, 倉庫: false }; });
  }

  function 直す(r) {
    return {
      いつ: r.itsu,
      種類: r.shurui,
      見出し: r.midashi,
      中身: r.nakami,
      ファイル名: r.file_name || '',
      クレジット: r.credit || 0,
    };
  }

  var api = {
    残す: 残す,
    読む: 読む,
    倉庫が使えるか: function () { return 倉庫が使えるか; },
    __仮置き: function () { return 仮置き; },
    __空にする: function () { 仮置き = []; 倉庫が使えるか = null; },
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.RirekiStore = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
