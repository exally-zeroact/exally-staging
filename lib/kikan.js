/* kikan.js — ★期間（何月分・◯日からの分）を決める 1本★
 * =============================================================================
 * ★借り物の正本（2026-08-27）★
 *   ★締め期間の決め方は Timeally が 正本★＝ここで 新しく考えない。
 *     正本 … timeally/lib/tc-calc.js の period(ym, closeDay)（頭SHA fa03839）
 *     試験 … timeally/tests/close-period.test.mjs（★締め日1〜31 × 月4種＝124通り★）
 *   ★同じ形のまま 置く★（読みやすく書き直さない）＝直った時に 見比べられなくなる。
 *   ★うちの試験でも 124通りを そのまま測る★（tests/kikan.test.mjs）。
 *   借りてよいのは ★道具・測り方・試験★ だけ（色・言葉・画面は 借りない）。
 *
 * ★司さんの言葉（2026-08-27）★
 *   「1人分と 何月分（1日からの分／10日からの分）って 分けて表示もさせれるん？」
 *   ⇒ ★「◯日からの分」＝締め期間★。「1日から」＝末日締め（締め日31）、
 *      「10日から」＝締め日9（前の月の10日〜その月の9日）。
 *
 * ★lib/periods.js とは 別物★（★2つ作った訳ではない★）
 *   periods.js … ★同じ月の中★を 1〜10／11〜20／21〜末 に割る（K2 報酬明細の期間）
 *   kikan.js   … ★月をまたぐ 締め期間★（前の月の◯日 〜 その月の◯日）
 *   ＝欲しい物が違う。どちらも 残す。
 *
 * 【利用】ブラウザ window.Kikan / Node require('./kikan.js')
 */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Kikan = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  function pad2(n) { return ('0' + n).slice(-2); }
  function daysInMonth(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); }

  /* ── ★ここから 借り物（timeally/lib/tc-calc.js の period）★ ────────
     closeDay = 31 は「末日締め」。それ以外は 前月(closeDay+1) 〜 当月(closeDay)。
     2月30日のような日は その月の末日に丸める（存在しない日で穴を開けない）。 */
  function period(ym, closeDay) {
    var y = +ym.slice(0, 4), m = +ym.slice(5, 7);
    var cd = Number(closeDay) || 31;
    if (cd >= 31) {
      return { ym: ym, from: ym + '-01', to: ym + '-' + pad2(daysInMonth(y, m)) };
    }
    var pm = m === 1 ? 12 : m - 1, py = m === 1 ? y - 1 : y;
    var fromD = Math.min(cd + 1, daysInMonth(py, pm));
    var toD = Math.min(cd, daysInMonth(y, m));
    return {
      ym: ym,
      from: py + '-' + pad2(pm) + '-' + pad2(fromD),
      to: y + '-' + pad2(m) + '-' + pad2(toD),
    };
  }
  /* ── ★ここまで 借り物★ ──────────────────────────────────────── */

  /** ★「◯日から」を 締め日に直す★（客が言うのは 始まりの日・機械が持つのは 締め日）
   *  1日から → 末日締め（31）／10日から → 締め日9 */
  function 始まりの日から締め日(始まりの日) {
    var d = Math.floor(Number(始まりの日));
    if (!(d >= 1 && d <= 31)) return null;
    return d === 1 ? 31 : d - 1;
  }

  /** ★何月分・◯日からの分★ → { from, to, 言い方 } */
  function 期間を決める(ym, 始まりの日) {
    if (!/^\d{4}-\d{2}$/.test(String(ym || ''))) return null;
    var cd = 始まりの日から締め日(始まりの日 === undefined || 始まりの日 === null ? 1 : 始まりの日);
    if (cd === null) return null;
    var p = period(ym, cd);
    var 月 = String(+ym.slice(5, 7));
    /* ★年をまたぐ時は 年を出す★＝「1月分（12/10〜1/9）」だと どの年の12月か 分からない */
    var またぐ = p.from.slice(0, 4) !== p.to.slice(0, 4);
    return {
      ym: ym, from: p.from, to: p.to, 締め日: cd,
      /* ★客に見せる言い方は ここ1か所で作る★（画面に散らさない） */
      言い方: 月 + '月分（' + 日の字(p.from, またぐ) + '〜' + 日の字(p.to, またぐ) + '）',
    };
  }

  /** 客に見せる日付の字（M/D。年をまたぐ時だけ YYYY/M/D） */
  function 日の字(ymd, 年を出す) {
    var t = String(ymd || '');
    var md = (+t.slice(5, 7)) + '/' + (+t.slice(8, 10));
    return 年を出す ? (t.slice(0, 4) + '/' + md) : md;
  }

  /* ── 日付（Excelの通し番号 46023 は 2026-01-01） ───────────────── */
  /** ★通し番号を 日付に直す★（★これを持たないと 期間で切れない★）
   *  ★1900年の うるう年の嘘★（実測 2026-08-27）
   *    Excel は 1900-02-29 が 在る事にしている（本当は 無い日）。
   *    ・1〜59  … 起点は 1899-12-31（1 は 1900-01-01・59 は 1900-02-28）
   *    ・60     … ★Excelの嘘の日 1900-02-29★（そのまま返す＝黙って 1日 ずらさない）
   *    ・61以上 … 起点は 1899-12-30（普段の日付は 全部 こちら）
   *  ★よく使う式 1本（1899-12-30 起点）だけだと 1〜59 が 1日 ずれる★
   *  ・時刻ぶんは 切り捨てる（1日の途中でも その日として扱う） */
  function 通し番号を日付に(n) {
    var v = Number(n);
    if (!isFinite(v) || v < 1) return null;
    var i = Math.floor(v);
    if (i === 60) return '1900-02-29';
    var 起点 = (i < 60) ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    var d = new Date(起点 + i * 86400000);
    return d.getUTCFullYear() + '-' + pad2(d.getUTCMonth() + 1) + '-' + pad2(d.getUTCDate());
  }

  /** セルの値を 日付に直す（通し番号／'YYYY-MM-DD'／'YYYY/M/D' を受ける・駄目なら null） */
  function 日付に直す(v) {
    if (v === null || v === undefined || v === '') return null;
    if (typeof v === 'number') return 通し番号を日付に(v);
    var s = String(v).trim();
    var m = /^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/.exec(s);
    if (m) return m[1] + '-' + pad2(+m[2]) + '-' + pad2(+m[3]);
    if (/^\d+(\.\d+)?$/.test(s)) return 通し番号を日付に(+s);
    return null;
  }

  /** 期間の中か（両端を含む・'YYYY-MM-DD' の字くらべ＝時差で変わらない） */
  function 期間の中か(ymd, 期間) {
    if (!ymd || !期間) return false;
    return ymd >= 期間.from && ymd <= 期間.to;
  }

  return {
    period: period,                       /* ★借り物の名前は 変えない★ */
    始まりの日から締め日: 始まりの日から締め日,
    期間を決める: 期間を決める,
    通し番号を日付に: 通し番号を日付に,
    日付に直す: 日付に直す,
    期間の中か: 期間の中か,
    日の字: 日の字,
  };
});
