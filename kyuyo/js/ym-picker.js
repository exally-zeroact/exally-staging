/* ym-picker.js — 対象月の選択を「年」「月」の2つの選択肢にする（iPhoneでも同じに動く）
 *
 * なぜ必要か（2026-08-04・司さんの実機で判明）:
 *   月の入力に <input type="month"> を使っていた。★iOS Safari は type="month" を持っていない。★
 *   持っていない type は ただの文字入力になるので、月が選べない（何を入れればいいか分からない）。
 *   端末や版で挙動が変わる物に頼らない＝★どの端末でも同じに動く物（select）に置き換える。★
 *
 * 使い方（HTML側）:
 *   <input type="hidden" data-ym class="finput scr-month" value="2026-08">
 *   → この部品が、その直前に「年」「月」の select を差し込む。
 *   ★hidden の input はそのまま残す＝既存のコードは今までどおり
 *     `el.value` を読む／書く、`change` を待つ、で動く（呼び出し側を1行も変えない）。
 *
 * 差し込みは MutationObserver で自動。あとから描き直される画面（賞与・随時改定など）でも
 * ★人が呼び忘れる余地を作らない。★
 *
 * 守り: tests/ios-unsupported.test.mjs が type="month"/week/datetime-local の再発を赤にする。
 */
(function (global) {
  'use strict';
  var doc = global.document;
  if (!doc) return;

  var MARK = 'ymPicked';

  function pad2(n) { return ('0' + n).slice(-2); }
  function parseYm(v) {
    var m = /^(\d{4})-(\d{2})$/.exec(String(v || ''));
    if (!m) return null;
    return { y: +m[1], m: +m[2] };
  }
  function thisYm() { var d = new Date(); return { y: d.getFullYear(), m: d.getMonth() + 1 }; }

  function yearsFor(cur) {
    var now = thisYm().y;
    var lo = Math.min(now - 6, cur - 1), hi = Math.max(now + 1, cur + 1);
    var out = [];
    for (var y = lo; y <= hi; y++) out.push(y);
    return out;
  }

  function enhance(input) {
    if (!input || input.dataset[MARK]) return;
    input.dataset[MARK] = '1';

    var cur = parseYm(input.value) || thisYm();
    var wrap = doc.createElement('span');
    wrap.className = 'ym-pick';
    var ys = doc.createElement('select'); ys.className = 'finput ym-y'; ys.setAttribute('aria-label', '年');
    var ms = doc.createElement('select'); ms.className = 'finput ym-m'; ms.setAttribute('aria-label', '月');
    yearsFor(cur.y).forEach(function (y) { var o = doc.createElement('option'); o.value = String(y); o.textContent = y + '年'; ys.appendChild(o); });
    for (var i = 1; i <= 12; i++) { var o = doc.createElement('option'); o.value = pad2(i); o.textContent = i + '月'; ms.appendChild(o); }
    wrap.appendChild(ys); wrap.appendChild(ms);
    if (input.parentNode) input.parentNode.insertBefore(wrap, input);

    var value = input.value;
    function paint() {
      var p = parseYm(value) || thisYm();
      // 収録していない年が来たら足す（過去の明細を開いた時に選べなくならないように）
      if (!Array.prototype.some.call(ys.options, function (o) { return o.value === String(p.y); })) {
        var o = doc.createElement('option'); o.value = String(p.y); o.textContent = p.y + '年';
        ys.insertBefore(o, ys.firstChild);
      }
      ys.value = String(p.y); ms.value = pad2(p.m);
    }
    // ★既存コードの `el.value = '2026-08'` で select も追随させる（呼び出し側を変えないため）
    try {
      Object.defineProperty(input, 'value', {
        configurable: true,
        get: function () { return value; },
        set: function (v) { value = String(v == null ? '' : v); paint(); },
      });
    } catch (e) { /* 定義できない環境では select 側の操作だけ効く */ }
    paint();

    function onPick() {
      value = ys.value + '-' + ms.value;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    ys.addEventListener('change', onPick);
    ms.addEventListener('change', onPick);
  }

  function scan(root) {
    var host = root || doc;
    if (!host.querySelectorAll) return;
    Array.prototype.forEach.call(host.querySelectorAll('input[data-ym]'), enhance);
  }

  // 最初の1回＋以後の描き直しを自動で拾う（人が呼び忘れる余地を作らない）
  function start() {
    scan(doc);
    if (global.MutationObserver && doc.body) {
      new global.MutationObserver(function () { scan(doc); }).observe(doc.body, { childList: true, subtree: true });
    }
  }
  if (doc.readyState === 'loading') doc.addEventListener('DOMContentLoaded', start);
  else start();

  global.YmPicker = { scan: scan, enhance: enhance, parseYm: parseYm };
})(typeof window !== 'undefined' ? window : globalThis);
