/* grid-refedit.js — ★式を直せるようにする（カーソルの位置で決める）★
 *
 * なぜ必要か（2026-08-06・司さんが実機で踏んだ）:
 *   A2 に「=B1+30」と打ってしまい、B1 を A1 に直したかった。
 *   ところが A1 を触ると ★式の末尾に A1 が足されて「=B1+30A1」★ になった。
 *   触るたびに悪くなり、★B1 を直す方法が無い★＝式を書き間違えたら終わり。
 *
 *   原因は「末尾がセル参照なら置換、でなければ末尾に追加」という作りだった。
 *   「=B1+30」は末尾が "30"（参照ではない）ので、必ず末尾に追加されていた。
 *
 * 直し方（Excelと同じ考え方）:
 *   ① ★カーソルが参照の上（または端）にある → その参照を置き換える★
 *      → 数式バーで B1 を触ってからセルを選べば、B1 だけが変わる
 *   ② カーソルの直前が = + - * / ( , : などの「参照を置ける場所」→ そこに差し込む
 *      → 今までどおり「=」→セル→「+」→セル と組み立てられる
 *   ③ どちらでもない（例：数字の直後）→ ★何もしない★
 *      → 勝手に式を壊さない。呼ぶ側が「どうすればいいか」を伝える
 *
 * 触らないこと:
 *   ・"文字列" の中にある A1 は参照ではない（置き換えない）
 *   ・SUM( や LOG10( のような関数名は参照ではない（置き換えない）
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.GridRefEdit = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* 参照を置ける場所（この直後なら差し込んでよい） */
  var OPENERS = '=＝+-*/^(,;:<>&%';

  /* "…" の中（文字列リテラル）の位置に印を付ける。中の A1 は参照ではない。 */
  function quotedMask(v) {
    var mask = new Array(v.length), inQ = false;
    for (var i = 0; i < v.length; i++) {
      var ch = v.charAt(i);
      if (ch === '"') { inQ = !inQ; mask[i] = true; continue; }
      mask[i] = inQ;
    }
    return mask;
  }

  /* 式の中の「本物のセル参照」を全部拾う（A1 / $A$1 / A1:B2） */
  function findRefs(v) {
    var RE = /\$?[A-Za-z]{1,3}\$?[0-9]{1,7}(?::\$?[A-Za-z]{1,3}\$?[0-9]{1,7})?/g;
    var mask = quotedMask(v), out = [], m;
    RE.lastIndex = 0;
    while ((m = RE.exec(v))) {
      var s = m.index, e = s + m[0].length;
      if (mask[s]) continue;                               // 文字列の中
      var prev = s > 0 ? v.charAt(s - 1) : '';
      var next = e < v.length ? v.charAt(e) : '';
      if (/[A-Za-z0-9_$]/.test(prev)) continue;            // 名前の途中（LOG10 の 10 など）
      if (next === '(') continue;                          // 関数名（LOG10( など）
      out.push({ s: s, e: e, text: m[0] });
    }
    return out;
  }

  /* 本体。返り: { v, pos, ok, why }  ok=false なら式は1文字も変えない */
  function refEditAt(value, cursor, addr) {
    var v = String(value === null || value === undefined ? '' : value);
    var pos = (typeof cursor === 'number' && cursor >= 0 && cursor <= v.length) ? cursor : v.length;
    var a = String(addr || '');
    if (!a) return { v: v, pos: pos, ok: false, why: 'no-addr' };

    /* ① カーソルが参照の上／端にある → 置き換える */
    var refs = findRefs(v);
    for (var i = 0; i < refs.length; i++) {
      if (pos >= refs[i].s && pos <= refs[i].e) {
        return {
          v: v.slice(0, refs[i].s) + a + v.slice(refs[i].e),
          pos: refs[i].s + a.length,
          ok: true, why: 'replace',
        };
      }
    }

    /* ② 直前が「参照を置ける場所」→ そこに差し込む */
    var before = v.slice(0, pos).replace(/\s+$/, '');
    var ch = before.charAt(before.length - 1);
    if (before === '' || OPENERS.indexOf(ch) >= 0) {
      return { v: v.slice(0, pos) + a + v.slice(pos), pos: pos + a.length, ok: true, why: 'insert' };
    }

    /* ③ それ以外 → ★何もしない★（式を壊さない） */
    return { v: v, pos: pos, ok: false, why: 'not-a-ref-position' };
  }

  /* ★F4＝$の付け方を1つ進める（カーソルの下の参照だけ）★
   *  ★実Excelで測った4つの形（2026-08-18・16.0.20228）★
   *      絶対      =A1+B2 → =$A$1+$B$2
   *      行だけ絶対 =A1+B2 → =A$1+B$2
   *      列だけ絶対 =A1+B2 → =$A1+$B2
   *      相対      =A1+B2 → =A1+B2
   *  ★回る順番だけは機械で測れていない★（F4 は COM から押せない）。
   *    Excel の標準の順（相対 → 絶対 → 行だけ → 列だけ → 相対）で作ってある。
   *    ★司さんが1回 押して確かめる物★として報告に載せる。 */
  function cycleAbs(ref) {
    var m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/.exec(String(ref));
    if (!m) return null;
    var col = m[2], row = m[4];
    var state = (m[1] ? 2 : 0) + (m[3] ? 1 : 0);   // 0:A1 3:$A$1 1:A$1 2:$A1
    var next = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
    return (next & 2 ? '$' : '') + col + (next & 1 ? '$' : '') + row;
  }

  /** カーソルの位置に在る参照の $ を1つ進める。
   *  無ければ ok:false（★式を1文字も変えない★） */
  function toggleAbsAt(v, pos) {
    var refs = findRefs(String(v));
    for (var i = 0; i < refs.length; i++) {
      if (pos >= refs[i].s && pos <= refs[i].e) {
        var cur = String(v).slice(refs[i].s, refs[i].e);
        /* A1:B2 の形は 両方を進める（Excelと同じ） */
        var parts = cur.split(':');
        var done = [];
        for (var k = 0; k < parts.length; k++) {
          var c = cycleAbs(parts[k]);
          if (!c) return { v: v, pos: pos, ok: false, why: 'not-a-ref' };
          done.push(c);
        }
        var rep = done.join(':');
        return {
          v: String(v).slice(0, refs[i].s) + rep + String(v).slice(refs[i].e),
          pos: refs[i].s + rep.length, ok: true, why: 'toggled',
        };
      }
    }
    return { v: v, pos: pos, ok: false, why: 'not-a-ref-position' };
  }

  return { refEditAt: refEditAt, findRefs: findRefs, cycleAbs: cycleAbs, toggleAbsAt: toggleAbsAt };
});
