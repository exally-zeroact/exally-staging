/* rireki.js — ★6 履歴＝「見る場所」★（純関数・0円）
 *
 *  ★決まり（司さん 2026-08-25／指示役 08-26）★
 *    ・★履歴は 見る場所★（探す入口にしない＝ここから何かを始めさせない）
 *    ・★置き場は 倉庫（会社の物）★／★客のブックにタブを1つも足さない★
 *      （競合は 客のブックに「Claude Log」タブを作る。★真似しない★）
 *    ・カテゴリ … 関数／自動化（レシピ）／直した所／診断／取り込み・書き出し／その他
 *    ・★今年は「8月25日 14:03」／去年は「2025年8月25日」★
 *    ・中身を載せる（何を聞いて／何をして／★どのセルが どう変わったか★）
 *    ・★ファイル名を長押しでコピー → AIタブで聞く★（携帯・PCの ふつうの仕組みのまま）
 *
 *  ★別の入り口（AIタブ／履歴ページ）から 同じに見える★＝
 *    ★一覧にする() を 1つだけ持つ★。画面は それを呼ぶだけ。
 */
(function (root) {
  'use strict';

  var 種類の順 = ['関数', '自動化', '直した所', '診断', '取り込み・書き出し', 'その他'];

  /** ★いつを 客の言い方に直す★（今年は 年を書かない・去年からは 年を書く） */
  function 日付の字(いつ, 今) {
    var d = (いつ instanceof Date) ? いつ : new Date(いつ);
    if (isNaN(d.getTime())) return '';
    var n = (今 instanceof Date) ? 今 : (今 ? new Date(今) : new Date());
    var 月 = d.getMonth() + 1, 日 = d.getDate();
    if (d.getFullYear() === n.getFullYear()) {
      var 時 = ('0' + d.getHours()).slice(-2), 分 = ('0' + d.getMinutes()).slice(-2);
      return 月 + '月' + 日 + '日 ' + 時 + ':' + 分;
    }
    return d.getFullYear() + '年' + 月 + '月' + 日 + '日';
  }

  /** ★種類を 決めた6つに寄せる★（知らない物は その他。★捨てない★） */
  function 種類に寄せる(s) {
    var t = String(s || '');
    return 種類の順.indexOf(t) >= 0 ? t : 'その他';
  }

  /**
   * ★一覧にする（★入り口が違っても ここを通る＝同じに見える★）★
   * @param {Array} 行たち [{いつ, 種類, 見出し, 中身, ファイル名, クレジット}]
   * @param {{今:Date, 種類:string, 上限:number}} opt
   * @returns {{件数:number, 出す:Array, 種類ごと:Object, 絞り:string, 出していない:number}}
   */
  function 一覧にする(行たち, opt) {
    opt = opt || {};
    var 今 = opt.今 || new Date();
    var 上限 = opt.上限 || 200;
    var 全部 = (行たち || []).map(function (r) {
      return {
        いつ: r.いつ,
        日付: 日付の字(r.いつ, 今),
        種類: 種類に寄せる(r.種類),
        見出し: String(r.見出し || ''),
        中身: r.中身 || null,
        ファイル名: r.ファイル名 ? String(r.ファイル名) : '',
        クレジット: Number(r.クレジット || 0),
      };
    });
    /* ★新しい物が 上★（見る場所なので 迷わせない） */
    全部.sort(function (a, b) { return new Date(b.いつ) - new Date(a.いつ); });
    var 種類ごと = {};
    for (var i = 0; i < 種類の順.length; i++) 種類ごと[種類の順[i]] = 0;
    for (var j = 0; j < 全部.length; j++) 種類ごと[全部[j].種類]++;
    var 絞った = opt.種類 ? 全部.filter(function (x) { return x.種類 === opt.種類; }) : 全部;
    var 出す = 絞った.slice(0, 上限);
    return {
      件数: 絞った.length,
      出す: 出す,
      種類ごと: 種類ごと,
      絞り: opt.種類 || '',
      /* ★多い時は 何件 出していないかを 必ず言う（黙って切らない）★ */
      出していない: Math.max(0, 絞った.length - 出す.length),
    };
  }

  /** ★1件の中身を 客の言葉にする★（どのセルが どう変わったか まで） */
  function 中身の字(件) {
    if (!件) return '';
    var c = 件.中身 || {};
    var 出 = [];
    if (c.聞いた) 出.push('聞いた事：' + c.聞いた);
    if (c.やった) 出.push('やった事：' + c.やった);
    if (c.何か所 !== undefined && c.何か所 !== null) 出.push('直した所：' + c.何か所 + 'か所');
    if (c.どこ && c.どこ.length) {
      var 先 = c.どこ.slice(0, 5).map(function (x) {
        return x.場所 + (x.前 !== undefined ? '（' + 見せる(x.前) + ' → ' + 見せる(x.後) + '）' : '');
      });
      出.push(先.join(' / ') + (c.どこ.length > 5 ? ' ほか ' + (c.どこ.length - 5) + 'か所' : ''));
    }
    if (件.クレジット > 0) 出.push('使ったクレジット：' + 件.クレジット);
    else 出.push('AIは使っていません');
    return 出.join('\n');
  }
  function 見せる(v) {
    if (v === undefined || v === null || v === '') return '（空）';
    return String(v);
  }

  var api = {
    種類の順: 種類の順,
    日付の字: 日付の字,
    種類に寄せる: 種類に寄せる,
    一覧にする: 一覧にする,
    中身の字: 中身の字,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.Rireki = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
