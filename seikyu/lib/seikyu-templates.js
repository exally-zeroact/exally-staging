/* seikyu-templates.js — ★様式（テンプレ）＝見た目と初期の列だけ★
 * ==============================================================================
 * 代行請求アプリと同じ考え方（invoice-pdf.js:324/338 の pdfDesign = classic / elegant）。
 *   classic … 緑の帯で罫線のしっかりした、いつもの請求書
 *   elegant … 罫線を減らして字で見せる、落ち着いた請求書
 *
 * ★テンプレが決めてよいのは「見た目」と「最初に並ぶ列」だけ★
 *   ・金額・消費税・合計は ★1円もテンプレに依らない★（totalsOf がそれを守る）
 *   ・列は最初の並びを配るだけ。あとは会社が足す・消す・幅を変える
 *
 * ★既定の7列は、ここに置いた「std1 の初期値」であって、紙に焼き付いた物ではない★
 *
 * 【利用】ブラウザ window.SeikyuTemplates ／ Node require('./seikyu-templates.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./seikyu-tax.js'), require('./seikyu-cols.js'));
  } else {
    root.SeikyuTemplates = factory(root.SeikyuTax, root.SeikyuCols);
  }
})(typeof self !== 'undefined' ? self : this, function (TAX, COLS) {
  'use strict';
  if (!TAX || !TAX.compute) throw new Error('seikyu-tax.js を先に読んでください');
  if (!COLS) throw new Error('seikyu-cols.js を先に読んでください');

  /* 色は★直hex★（このリポジトリの現行ルール）。#1A4A2E は使わない。 */
  var TEMPLATES = {
    std1: {
      id: 'std1',
      label: 'いつもの（緑の帯・罫線あり）',
      note: '代行請求で毎日 刷っている形に近い、いちばん普通の請求書です。',
      cols: {
        items: ['#', '品名・内容', '数量', '単位', '単価', '金額', '税率'],
        widths: { '#': 28, '品名・内容': 220, '数量': 56, '単位': 44, '単価': 80, '金額': 100, '税率': 56 },
        aligns: {},
      },
      theme: {
        ink: '#24422F',       // 本文
        sub: '#5C7E6C',       // 補足
        line: '#D4EAE0',      // 罫線
        accent: '#2E7D54',    // 見出し
        band: '#F0FAF4',      // 帯の地
        headBg: '#F0FAF4',    // 表の見出しの地
        headInk: '#2E7D54',
        grandInk: '#2E7D54',
        rule: 'all',          // 罫線 = 全部引く
        titleSpacing: '.32em',
      },
    },
    elegant: {
      id: 'elegant',
      label: 'すっきり（罫線ひかえめ・字で見せる）',
      note: '罫線を減らして、字の大きさと余白で読ませる形です。',
      cols: {
        items: ['#', '品名・内容', '数量', '単位', '単価', '金額', '税率'],
        widths: { '#': 28, '品名・内容': 240, '数量': 56, '単位': 44, '単価': 84, '金額': 108, '税率': 56 },
        aligns: {},
      },
      theme: {
        ink: '#24422F',
        sub: '#7AA08C',
        line: '#E4F1EA',
        accent: '#3D9E72',
        band: '#FFFFFF',
        headBg: '#FFFFFF',
        headInk: '#3D9E72',
        grandInk: '#3D9E72',
        rule: 'rows',         // 罫線 = 横線だけ
        titleSpacing: '.5em',
      },
    },
  };

  var DEFAULT_ID = 'std1';

  function get(id) {
    var t = TEMPLATES[String(id || '')] || null;
    return t ? clone(t) : null;
  }
  function getOrDefault(id) { return get(id) || get(DEFAULT_ID); }
  function list() {
    return Object.keys(TEMPLATES).map(function (k) {
      var t = TEMPLATES[k];
      return { id: t.id, label: t.label, note: t.note };
    });
  }
  function clone(v) { return JSON.parse(JSON.stringify(v)); }

  /** その1通で使う列。会社が決めた物（inv.data.cols）が最優先、無ければテンプレの初期値。 */
  function colsOf(inv) {
    var d = (inv && inv.data) || {};
    var own = d.cols;
    if (own && Array.isArray(own.items) && own.items.length) return COLS.normalizeSpec(own);
    return COLS.normalizeSpec(getOrDefault(inv && inv.template_id).cols);
  }

  /**
   * ★金額は様式に依らない★
   *   ここは template_id を1度も見ない。見た瞬間に「見た目で金額が変わる」バグになる。
   *   totalsOf({ inv, lines }) → seikyu-tax.compute の返り
   */
  function totalsOf(o) {
    o = o || {};
    var inv = o.inv || {};
    return TAX.compute({
      lines: o.lines || inv.lines || [],
      taxMode: inv.tax_mode,
      rounding: inv.rounding,
    });
  }

  return {
    TEMPLATES: TEMPLATES, DEFAULT_ID: DEFAULT_ID,
    get: get, getOrDefault: getOrDefault, list: list,
    colsOf: colsOf, totalsOf: totalsOf,
  };
});
