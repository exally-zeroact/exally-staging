/* grid-valid.js — ★入力の決まり（データの入力規則）★（純関数）
 *
 *  ★実Excel 16.0.20228 を COM で押して測った（2026-08-21）★
 *
 *  ① 一覧から選ぶ（xlValidateList=3）
 *       中身「りんご,みかん,ぶどう」／★ドロップダウンを出す=True★／
 *       ★空白は見ない(IgnoreBlank)=True★／エラーを出す=True／出し方=停止(1)
 *  ② 整数（xlValidateWholeNumber=1）1〜10
 *       5=合う ／ 99=合わない ／ ★5.5=合わない（小数は整数の決まりに反する）★
 *  ③ ★空にすると「合っている」扱い★（IgnoreBlank=True のため）
 *
 *  ★一番 大事な実測（ここを間違えると嘘になる）★
 *    ④ ★決まりに反する値でも 打つ以外の道（貼り付け・読み込み・並べ替え）では 入ってしまう★。
 *       COM で 'すいか' を入れたら ★そのまま入り★、Validation.Value だけ False になった。
 *    ⑤ ★先に値を入れてから 決まりを足しても 値は消えない★（合っていないまま 残る）。
 *    ⇒ ★入力の決まりは「打つ時に止める」仕掛けであって「消す仕掛け」ではない★。
 *
 *  ★うちの方が良い所（違いとして残す）★
 *    Excel は 入ってしまった「合っていない値」を ★黙って抱える★（○を付けるまで見えない）。
 *    うちは ★合っていないセルが何個 在るかを 数えて出す★＝★黙って通さない★。
 */
(function (root) {
  'use strict';

  /** 一覧の中身を切る（Excelはカンマ区切り。前後の空白は落とす） */
  function parseList(src) {
    return String(src === undefined || src === null ? '' : src)
      .split(',').map(function (s) { return s.trim(); })
      .filter(function (s) { return s !== ''; });
  }

  function isBlank(text) {
    return text === undefined || text === null || String(text).trim() === '';
  }

  /**
   * 打った字が 決まりに合っているか
   * @returns {{ok:boolean, why:string}}  ok=false の時 why に ★人の言葉の理由★
   */
  function check(rule, text) {
    if (!rule) return { ok: true, why: '' };
    /* ★空は合っている扱い（実測 IgnoreBlank=True）★ */
    if (isBlank(text)) return { ok: true, why: '' };
    var s = String(text).trim();

    if (rule.kind === 'list') {
      var items = rule.items || [];
      for (var i = 0; i < items.length; i++) if (items[i] === s) return { ok: true, why: '' };
      return { ok: false, why: 'この列は ' + items.join('／') + ' から選びます' };
    }

    if (rule.kind === 'int' || rule.kind === 'num') {
      var n = Number(s.replace(/,/g, ''));
      if (s === '' || isNaN(n)) return { ok: false, why: '数を入れてください' };
      /* ★整数の決まりに 5.5 は合わない（実測）★ */
      if (rule.kind === 'int' && Math.floor(n) !== n) {
        return { ok: false, why: '整数を入れてください（小数は入りません）' };
      }
      var 下 = (rule.min === undefined || rule.min === null || rule.min === '') ? null : Number(rule.min);
      var 上 = (rule.max === undefined || rule.max === null || rule.max === '') ? null : Number(rule.max);
      if (下 !== null && n < 下) return { ok: false, why: 下 + ' 以上を入れてください' };
      if (上 !== null && n > 上) return { ok: false, why: 上 + ' 以下を入れてください' };
      return { ok: true, why: '' };
    }

    return { ok: true, why: '' };
  }

  /** 決まりを 人の言葉で1行にする（画面に出して 確かめられるように） */
  function describe(rule) {
    if (!rule) return '';
    if (rule.kind === 'list') return '一覧から選ぶ：' + (rule.items || []).join('／');
    var 下 = (rule.min === undefined || rule.min === null || rule.min === '') ? null : rule.min;
    var 上 = (rule.max === undefined || rule.max === null || rule.max === '') ? null : rule.max;
    var 名 = (rule.kind === 'int') ? '整数' : '数';
    if (下 !== null && 上 !== null) return 名 + '：' + 下 + '〜' + 上;
    if (下 !== null) return 名 + '：' + 下 + ' 以上';
    if (上 !== null) return 名 + '：' + 上 + ' 以下';
    return 名;
  }

  function key(r, c) { return r + ',' + c; }

  /** 決まりを 範囲に付ける（rules は書き換えず 新しい物を返す） */
  function setRange(rules, r1, c1, r2, c2, rule) {
    var out = {}, k;
    for (k in rules) out[k] = rules[k];
    for (var r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) {
      for (var c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        if (rule) out[key(r, c)] = rule; else delete out[key(r, c)];
      }
    }
    return out;
  }

  /**
   * ★合っていないセルを 数える★（Excelには無い＝黙って通さないため）
   *  ★打っていない道（貼り付け・読み込み・並べ替え）で 入った値を 見つける★
   * @returns {{count:number, cells:Array<{r:number,c:number,text:string,why:string}>}}
   */
  function countBad(data, rules, shownOf) {
    var 出す = shownOf || function (cell) {
      if (!cell) return '';
      var v = (cell.d !== undefined && cell.d !== null && cell.d !== '') ? cell.d : cell.v;
      return (v === undefined || v === null) ? '' : String(v);
    };
    var cells = [];
    for (var k in (rules || {})) {
      var p = k.split(','), r = +p[0], c = +p[1];
      var t = 出す(data ? data[k] : null);
      var res = check(rules[k], t);
      if (!res.ok) cells.push({ r: r, c: c, text: t, why: res.why });
    }
    cells.sort(function (a, b) { return (a.r - b.r) || (a.c - b.c); });
    return { count: cells.length, cells: cells };
  }

  var api = {
    parseList: parseList, check: check, describe: describe,
    setRange: setRange, countBad: countBad,
  };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.GridValid = api;
})(typeof globalThis !== 'undefined' ? globalThis : this);
