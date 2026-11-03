/* mochikomi.js — ★持ち込んだ ファイルの 画面★（mochikomi.html から 使う）
 *
 *  ★★なぜ 別ページか（司さん 2026-09-21・ア）★★
 *    「★ユーザーが 持ってきた ファイルは 別ページで 保存★
 *      （どんな 関数や マクロが 組まれてるか 説明、★ドロップダウンで 詳しく★）」
 *
 *  ★★この 画面が する 事／しない 事★★
 *    する  ... 開く ／ ★中身を 数えて 言う★ ／ ★そのまま 返す★
 *    しない ... ★直す★（1マスも 打てません）
 *      ⇒★「持ち込んだ 物を 触らない」事を 画面の 作りで 保証します★
 *      ⇒直すのは `book.html`。★2つの 道を 分けるのが この 画面の 役目★
 *
 *  ★★作り直して いません★★（`find-existing` を 走らせて 数えました）
 *    開く     ... `js/book-open.js` の `BookOpen.openFile`（★お客さんの 道と 同じ★）
 *    数える   ... `lib/hon-no-nakami.js`
 *    関数名   ... `lib/excel-version.js` の `functionsIn`
 *    マクロ   ... `lib/vba.js` ／ `lib/vba-mikata.js`
 *    返す     ... `js/book-open.js` の `BookOpen.saveOpened` ＋ `js/file-out.js`
 *    ⇒★1つも 自前で 解きません★
 *
 *  ★★「1バイトも 変えない」と 言うなら 数えます★★
 *    返した 後に ★元の 大きさと 出した 大きさ★を 並べて 出します。
 *    ⇒★言うだけに しない★（記憶「『紙が 在る』は 押されてる・分かる・確かめられる の 3つとも 別」）
 */
(function (global) {
  'use strict';

  var 今 = null;           /* 開いた 本（BookOpen.openFile の 返り） */
  var 元の大きさ = 0;

  function $(id) { return document.getElementById(id); }
  function 字(el, s) { if (el) el.textContent = s; }

  function 言う(s, 色) {
    var e = $('shirase');
    if (!e) return;
    e.textContent = s;
    e.className = 'shirase' + (色 ? ' ' + 色 : '');
    e.hidden = !s;
  }

  /** ★ドロップダウンを 1つ 作る★（★中身は 台が 出した 行を そのまま★） */
  function 畳み(見出し, 行たち, 注) {
    var d = document.createElement('details');
    var s = document.createElement('summary');
    s.textContent = 見出し + '（' + 行たち.length + '件）';
    d.appendChild(s);
    if (注) {
      var p = document.createElement('p');
      p.className = 'chu';
      p.textContent = 注;
      d.appendChild(p);
    }
    if (行たち.length) {
      var ul = document.createElement('ul');
      for (var i = 0; i < 行たち.length; i++) {
        var li = document.createElement('li');
        li.textContent = 行たち[i];
        ul.appendChild(li);
      }
      d.appendChild(ul);
    }
    return d;
  }

  function 出す(res, file) {
    今 = res;
    元の大きさ = file.size;

    字($('fname'), file.name);
    字($('fmeta'), (res.kind || '') + ' ／ ' + file.size.toLocaleString() + ' バイト ／ 板 '
      + (res.sheets || []).length + '枚');

    /* ★マクロは 呼ぶ側が 読んで 渡す★（台は 自分で 開かない） */
    var 読み = null, 見立て = null;
    if (res.hasVba && res.マクロ) {
      読み = res.マクロ.読み || res.マクロ;
      見立て = res.マクロ.見立て || null;
      if (!見立て && global.VbaMikata && 読み && 読み.ok) {
        try { 見立て = global.VbaMikata.見立てる(読み.モジュール); } catch (e) { 見立て = null; }
      }
    }

    var 数 = global.HonNoNakami.本の中身を数える({
      形: res.kind,
      sheets: res.sheets,
      関数を拾う: (global.ExcelVersion && global.ExcelVersion.functionsIn) || null,
      マクロの読み: 読み,
      マクロの見立て: 見立て
    });

    字($('hitokoto'), global.HonNoNakami.一文(数));

    var 箱 = $('kuwashiku');
    箱.innerHTML = '';
    var 並び = global.HonNoNakami.詳しく(数);
    for (var i = 0; i < 並び.length; i++) {
      箱.appendChild(畳み(並び[i].見出し, 並び[i].行 || [], 並び[i].注 || ''));
    }

    $('kekka').hidden = false;
    $('hozon').disabled = false;
    言う('');
  }

  function 開く(file) {
    if (!file) return;
    言う('読み込んでいます...', '');
    $('hozon').disabled = true;
    $('kekka').hidden = true;
    global.BookOpen.openFile(file).then(function (res) {
      出す(res, file);
    }).catch(function (e) {
      言う('開けませんでした ... ' + (e && e.message ? e.message : e), 'akai');
    });
  }

  function 返す() {
    if (!今 || !global.BookOpen.isOpened()) { 言う('先に ファイルを 選んでください。', 'akai'); return; }
    var cur = global.BookOpen.current() || {};
    $('hozon').disabled = true;
    言う('書き出しています...', '');
    global.BookOpen.saveOpened(今.sheets).then(function (res) {
      var b = (res && res.bytes) ? res.bytes : res;
      var 出た = (b && b.length) || 0;
      return global.FileOut.deliver(b, cur.name).then(function () {
        /* ★「1バイトも 変えていません」を ★数で★ 出します★ */
        var 同じ = (出た === 元の大きさ);
        字($('hozon-kekka'),
          '書き出しました ... 元 ' + 元の大きさ.toLocaleString() + ' バイト ／ 出した '
          + 出た.toLocaleString() + ' バイト'
          + (同じ ? '（大きさは 同じです）' : '（大きさが 違います 差 '
            + (出た - 元の大きさ).toLocaleString() + ' バイト）'));
        $('hozon-kekka').hidden = false;
        言う('');
      });
    }).catch(function (e) {
      言う('書き出せませんでした ... ' + (e && e.message ? e.message : e), 'akai');
    }).then(function () { $('hozon').disabled = false; });
  }

  global.Mochikomi = {
    つなぐ: function () {
      var inp = $('finput');
      if (inp) {
        inp.addEventListener('change', function () {
          var f = this.files && this.files[0];
          this.value = '';          /* ★同じ ファイルを もう一度 選べる ように★ */
          開く(f);
        });
      }
      var b = $('erabu');
      if (b) b.addEventListener('click', function () { $('finput').click(); });
      var h = $('hozon');
      if (h) h.addEventListener('click', 返す);
    },
    _開く: 開く          /* ★試験から 呼ぶ 口★（画面の ボタンと 同じ 道） */
  };
}(typeof self !== 'undefined' ? self : this));
