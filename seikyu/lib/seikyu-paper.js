/* seikyu-paper.js — ★うちの紙（請求書）を刷る唯一の場所★
 * ==============================================================================
 * ★作法の一次情報 = daikou-seikyu-test/invoice-pdf.js★（毎日 113通 刷っている実物）
 *   466行  「御請求金額（枠なし＝アプリ統一）。ラベル＋大きい金額＋下に線」
 *   460行  「下記の通り御請求申し上げます。」（その上に「{月}月のご利用分です。」）
 *   156行  yen() = "¥" + toLocaleString      ← ★金額は ¥ 記号★
 *   408/413/428行  「請求日　」「No.　」「お支払期限　」 ← ★ラベルの後ろは全角スペース★
 *   180行  issueDateStr … 和暦（令和X年M月D日）／西暦（YYYY/M/D）を選べる
 *   571行  小計/消費税/合計（右下・★枠なし★・合計の上に線＋太字）
 *   720行  表の上に【…】の小さなキャプション
 *   755/760/785行  「このページの小計」「次ページへ続く →」「nページ目」
 *   420行  宛名は `会社名　御中`（全角スペース・下線は引かない）
 *
 * ★ここは「同じ作法で刷る」。真似ではない。★
 *
 * ★1つだけ、そのまま持ってこられない物＝色★
 *   invoice-pdf.js の MINT は今 ★#007AFF（青）★、本文は #0A5FD0（青）です（219〜233行）。
 *   代行請求は ★ダイコメの製品★ で、事務所の青に寄せた配色に変わっています。
 *   （同じ場所のコメントに「旧値 0.102,0.29,0.18 は 使わないと決めた濃い緑そのものだった」とある）
 *   請求書は Exally／Kyually の緑の家なので、★形は同じ・色はうちの緑★で刷ります。
 *   ここは勝手に決めず、報告で司さんに出します。青にするならこの1箇所を差し替えるだけです。
 *
 * ★税率の数字を1つも書かない★（区分は seikyu-tax.js が出した物を並べるだけ）
 * ★画面に依らない（DOMを1つも触らない）／時計も乱数も持たない★
 *
 * 【利用】ブラウザ window.SeikyuPaper ／ Node require('./seikyu-paper.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./seikyu-cols.js'));
  else root.SeikyuPaper = factory(root.SeikyuCols);
})(typeof self !== 'undefined' ? self : this, function (COLS) {
  'use strict';
  if (!COLS) throw new Error('seikyu-cols.js を先に読んでください');

  var TEMPLATE_ID = 'std1';

  /* 1枚に載る明細の行数。★ここを超えたら次の紙へ送る（黙って切らない）★
     1枚目は 宛名・挨拶・御請求金額 が乗るぶん少ない。 */
  var ROWS_FIRST = 12;
  var ROWS_REST = 24;

  /* 色は★直hex★。#1A4A2E は使わない。 */
  var THEME = {
    ink: '#24422F',       // 本文
    sub: '#5C7E6C',       // 補助文（挨拶・ラベル）
    line: '#D4EAE0',      // 行間の細い罫
    accent: '#2E7D54',    // 飾り線・見出し
    headBg: '#F0FAF4',    // 表の見出しの地
    headInk: '#2E7D54',
    grandInk: '#2E7D54',
    rule: 'rows',
    titleSpacing: '.32em',
    grandGo: 'ご',        // ご請求金額（税込）＝ classic系の言い方（invoice-pdf.js:957）
  };
  function themeOf(t) { return Object.assign({}, THEME, t || {}); }

  var DEFAULT_COLS = {
    items: ['#', '品名・内容', '数量', '単位', '単価', '金額', '税率'],
    widths: { '#': 28, '品名・内容': 220, '数量': 56, '単位': 44, '単価': 80, '金額': 100, '税率': 56 },
    aligns: {},
  };

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ★金額は ¥ 記号（invoice-pdf.js:156 と同じ作り）★
     数にならない物は 0 にしない（取れなかったを 0 と作り分ける）。 */
  function yen(v) {
    var n = Number(v);
    if (!Number.isFinite(n)) return '—';
    return '¥' + Math.round(n).toLocaleString('ja-JP');
  }
  /* 表の中の数（¥ を付けない・桁区切りだけ）＝ invoice-pdf.js の comma() と同じ */
  function comma(v) {
    if (v === undefined || v === null || v === '') return '';
    var n = Number(v);
    if (!Number.isFinite(n)) return esc(v);
    return n.toLocaleString('ja-JP');
  }
  function num(v) { return comma(v); }

  /* ★日付＝和暦か西暦（invoice-pdf.js:180 と同じ選び方）★
     era='reiwa' … 令和8年9月30日 ／ それ以外 … 2026/9/30
     読めない日付は空（勝手に今日を入れない）。 */
  function dateStr(ymd, era) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
    if (!m) return '';
    var y = +m[1], mo = +m[2], d = +m[3];
    if (era === 'reiwa') return '令和' + (y - 2018) + '年' + mo + '月' + d + '日';
    return y + '/' + mo + '/' + d;
  }
  /* 昔の紙との読み合わせ用（残す）。和暦なしの長い形。 */
  function jpDate(ymd) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
    if (!m) return '';
    return +m[1] + '年' + (+m[2]) + '月' + (+m[3]) + '日';
  }

  /* 敬称。★取引先マスタは hub が `keisho` で書いている★ ので両方を見る。 */
  function honorOf(p) {
    var h = (p && (p.honor || p.keisho)) || '';
    if (!h || h === '（なし）' || h === '(なし)' || h === 'なし') return '';
    return h;
  }

  /* 消費税のラベル。★率の数字は書かない＝区分から作る★
     1種類だけ … 消費税（10%）／内税なら 消費税（10%・内税）
     2種類以上 … 消費税（内訳は下に出す） */
  function taxLabel(tax, taxMode) {
    var by = (tax && tax.byRate) || [];
    var inTax = (taxMode === 'inclusive') ? '・内税' : '';
    if (by.length === 1) return '消費税（' + comma(by[0].pct) + '%' + inTax + '）';
    return '消費税' + (inTax ? '（内税）' : '');
  }

  function hasRole(items, role) {
    for (var i = 0; i < items.length; i++) if (COLS.roleOf(items[i]) === role) return true;
    return false;
  }

  /* 明細を紙ごとに分ける（★黙って切らない・次の紙へ送る★） */
  function paginate(lines, o) {
    var first = (o && o.rowsFirst) || ROWS_FIRST;
    var rest = (o && o.rowsRest) || ROWS_REST;
    if (!lines.length) return [[]];
    if (lines.length <= first) return [lines.slice()];
    var pages = [lines.slice(0, first)];
    for (var i = first; i < lines.length; i += rest) pages.push(lines.slice(i, i + rest));
    return pages;
  }

  /* ═══ 紙 ═══
   * build({ inv, tax, partner, org, cols, theme, era, page })
   *   era … 'reiwa' で和暦。既定は西暦（代行請求の既定 dateEra:'seireki' と同じ）
   */
  function build(o) {
    o = o || {};
    var inv = o.inv || {};
    var tax = o.tax || {};
    var p = o.partner || {};
    var g = o.org || {};
    var TH = themeOf(o.theme);
    var era = o.era || (inv.data && inv.data.dateEra) || 'seireki';

    var spec = COLS.normalizeSpec((o.cols && o.cols.items && o.cols.items.length) ? o.cols : DEFAULT_COLS);
    if (COLS.validate(spec.items).length) spec = COLS.normalizeSpec(DEFAULT_COLS);
    var colW = COLS.widthsOf(spec.items, spec.widths);

    var isQuote = (inv.doc_type === 'quote');
    var heading = isQuote ? '見　積　書' : '請　求　書';
    /* ★ラベルは様式で違う（実物を刷って確かめた）★
         classic系 … 「ご請求金額（税込）」（invoice-pdf.js:957）
         elegant系 … 「御請求金額（税込）」（invoice-pdf.js:479）
       どちらもうちの言葉。テーマが持つ。 */
    var go = TH.grandGo || 'ご';
    var grandLabel = go + (isQuote ? '見積金額（税込）' : '請求金額（税込）');
    var noLabel = 'No.　';
    var docTitle = (o.title || (heading.replace(/　/g, '') + (inv.no ? ' ' + inv.no : '')));

    var lines = Array.isArray(tax.lines) ? tax.lines : [];
    var byRate = Array.isArray(tax.byRate) ? tax.byRate : [];
    var exemptBase = (tax.exempt && Number(tax.exempt.base)) || 0;
    var pages = paginate(lines, o.page);
    var multi = pages.length > 1;

    /* 明細の見出し（items のとおり・その順） */
    var headHtml = spec.items.map(function (k, c) {
      return '<th class="c-col" style="width:' + colW[c].toFixed(4) + '%;text-align:' + COLS.alignOf(spec, k) + '">' + esc(k) + '</th>';
    }).join('');

    function rowsHtmlOf(pageLines, offset) {
      if (!pageLines.length) {
        return '<tr><td class="c-empty" colspan="' + spec.items.length + '">明細がまだ1行もありません</td></tr>';
      }
      return pageLines.map(function (ln, i) {
        return '<tr>' + spec.items.map(function (k) {
          var cell = COLS.cellOf(ln, k, offset + i);
          var al = COLS.alignOf(spec, k);
          var role = COLS.roleOf(k);
          var noWrap = (role === 'rate' || role === 'unit' || role === 'index');
          var body = cell.kind === 'money' ? (cell.text === '' ? '' : comma(cell.text))
            : cell.kind === 'num' ? (cell.text === '' ? '' : num(cell.text))
              : esc(cell.text);
          if (role === 'name' && ln.memo && !hasRole(spec.items, 'memo')) {
            body += '<span class="c-memo">' + esc(ln.memo) + '</span>';
          }
          return '<td class="c-col c-' + al + ((cell.kind === 'text' && !noWrap) ? ' c-wrap' : '')
            + (noWrap ? ' c-nowrap' : '') + '">' + body + '</td>';
        }).join('') + '</tr>';
      }).join('');
    }

    /* ── 頭（宛名・自社・日付） ★ラベルの後ろは全角スペース★ ── */
    function headBlock(pageIdx) {
      /* ★並びは実物と同じ 請求日 → No. → お支払期限（invoice-pdf.js:408→413→428）★
         ★番号は空でも欄を出す（「（未採番）」と書く）＝取れなかったを空欄にしない */
      var ds = dateStr(inv.issue_ymd, era);
      var meta = '<div class="meta-l">' + (isQuote ? '見積日　' : '請求日　') + (ds || '（未入力）') + '</div>';
      meta += '<div class="meta-l">' + esc(noLabel) + (esc(inv.no) || '（未採番）') + '</div>';
      if (inv.due_ymd) meta += '<div class="meta-l">お支払期限　' + dateStr(inv.due_ymd, era) + '</div>';

      return '<h1 class="ttl">' + heading + '</h1>'
        + '<div class="meta">' + meta + '</div>'
        + '<table class="party"><tbody><tr>'
        + '<td class="party-to">'
        + '<div class="to-name">' + (esc(p.name) || '（取引先が未選択）') + (honorOf(p) ? '　' + esc(honorOf(p)) : '') + '</div>'
        + (p.person ? '<div class="to-sub">' + esc(p.person) + '　様</div>' : '')
        + (p.zip ? '<div class="to-sub">〒' + esc(p.zip) + '</div>' : '')
        + (p.addr ? '<div class="to-sub">' + esc(p.addr) + '</div>' : '')
        + '</td>'
        + '<td class="party-from">'
        + '<div class="from-name">' + (esc(g.yago) || '（自社情報が未入力）') + '</div>'
        + (g.addr ? '<div class="from-sub">' + esc(g.addr) + '</div>' : '')
        + (g.tel ? '<div class="from-sub">TEL ' + esc(g.tel) + '</div>' : '')
        + (g.invoiceNo ? '<div class="from-sub">登録番号 ' + esc(g.invoiceNo) + '</div>' : '')
        + (g.sealDataUrl ? '<img class="seal" src="' + esc(g.sealDataUrl) + '" alt="">' : '')
        + '</td></tr></tbody></table>'
        + (multi ? '<div class="pageno">' + (pageIdx + 1) + 'ページ目</div>' : '');
    }

    /* ── 挨拶（★下記の通り御請求申し上げます。★） ── */
    function leadBlock() {
      var lead = (inv.data && inv.data.lead) || '';
      var mm = /^(\d{4})-(\d{2})-/.exec(String(inv.issue_ymd || ''));
      if (!lead && mm) lead = (+mm[2]) + '月分のご利用分です。';
      var greet = isQuote ? '下記の通り御見積申し上げます。' : '下記の通り御請求申し上げます。';
      return '<div class="lead">'
        + (lead ? '<div class="lead-l">' + esc(lead) + '</div>' : '')
        + '<div class="lead-l">' + greet + '</div>'
        + '</div>';
    }

    /* ── 御請求金額（★枠なし・ラベル＋大きい金額・下に線★） ── */
    function grandBlock() {
      return '<div class="grand">'
        + '<span class="grand-l">' + grandLabel + '</span>'
        + '<span class="grand-v">' + yen(tax.grandTotal) + '</span>'
        + '</div>';
    }

    /* ── 小計・消費税・合計（★枠なし・合計の上に線★） ── */
    function totalsBlock() {
      var rows = ''
        + '<tr><th>小計</th><td>' + yen(tax.subtotal) + '</td></tr>'
        + '<tr><th>' + taxLabel(tax, inv.tax_mode) + '</th><td>' + yen(tax.taxTotal) + '</td></tr>'
        + '<tr class="sums-g"><th>合計</th><td>' + yen(tax.grandTotal) + '</td></tr>';
      return '<table class="sums"><tbody>' + rows + '</tbody></table>';
    }

    /* ── （内訳）＝税率ごとの区分（適格請求書の要件）★枠で囲まない★ ── */
    function breakdownBlock() {
      var rows = byRate.map(function (b) {
        return '<tr><th>' + comma(b.pct) + '% 対象</th>'
          + '<td>' + yen(b.base) + '</td><td>' + yen(b.tax) + '</td></tr>';
      }).join('');
      if (exemptBase !== 0) {
        rows += '<tr><th>消費税の対象外</th><td>' + yen(exemptBase) + '</td><td>—</td></tr>';
      }
      if (!rows) rows = '<tr><td class="r-none" colspan="3">区分はまだありません</td></tr>';
      return '<div class="bd"><div class="bd-h">（内訳）</div>'
        + '<table class="rates"><thead><tr><th>区分</th><th>対象額</th><th>消費税</th></tr></thead>'
        + '<tbody>' + rows + '</tbody></table></div>';
    }

    /* ── 足元（★左＝お振込先／右＝小計・消費税・合計＋（内訳）★）
         invoice-pdf.js:979「合計フッター（左=振込先／右=小計・消費税・合計）」と同じ並び。
         ★2段組みは表で作る★（flex だと文が1文字ずつ縦に割れる） */
    function footerBlock() {
      var left = '';
      if (g.bank) left += '<div class="note"><div class="note-h">お振込先</div><div class="note-b">' + esc(g.bank).replace(/\n/g, '<br>') + '</div></div>';
      var memo = (inv.data && inv.data.memo) || '';
      if (memo) left += '<div class="note"><div class="note-h">備考</div><div class="note-b">' + esc(memo).replace(/\n/g, '<br>') + '</div></div>';
      return '<table class="foot"><tbody><tr>'
        + '<td class="foot-l">' + left + '</td>'
        + '<td class="foot-r">' + totalsBlock() + breakdownBlock() + '</td>'
        + '</tr></tbody></table>';
    }

    var subject = (inv.data && inv.data.subject) || '';
    var caption = subject || (inv.data && inv.data.tableTitle) || '';

    /* ── 紙を組む ── */
    var sheets = pages.map(function (pageLines, idx) {
      var last = (idx === pages.length - 1);
      var offset = pages.slice(0, idx).reduce(function (a, x) { return a + x.length; }, 0);
      var pageSum = pageLines.reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0);
      var body = ''
        + headBlock(idx)
        + (idx === 0 ? leadBlock() + grandBlock() : '')
        + (caption ? '<div class="cap">【' + esc(caption) + '】</div>' : '')
        + '<table class="items"><thead><tr>' + headHtml + '</tr></thead>'
        + '<tbody>' + rowsHtmlOf(pageLines, offset) + '</tbody></table>'
        + (last
          ? footerBlock()
          : '<div class="cont">'
            + '<div class="cont-l">このページの小計<span class="cont-v">' + yen(pageSum) + '</span></div>'
            + '<div class="cont-n">次ページへ続く →</div></div>');
      return '<div class="sheet">' + body + '</div>';
    }).join('');

    var html = ''
      + '<!DOCTYPE html>\n<html lang="ja"><head><meta charset="UTF-8">'
      + '<meta name="viewport" content="width=device-width, initial-scale=1">'
      + '<title>' + esc(docTitle) + '</title>'
      + '<style>' + css(TH) + '</style></head><body>'
      + sheets
      + '</body></html>';

    return {
      html: html, title: docTitle, templateId: inv.template_id || TEMPLATE_ID,
      cols: spec, colWidths: colW, pages: pages.length,
    };
  }

  /* ── 紙の見た目 ────────────────────────────────────────────
     ★文が入る所に flex/grid を使わない★（1文字ずつ縦に割れる事故を作らない）。
     ★枠で囲まない★（御請求金額・振込先・備考）＝うちの紙の作法。 */
  function css(t) {
    var TH = themeOf(t);
    var INK = TH.ink, SUB = TH.sub, LINE = TH.line, ACCENT = TH.accent;
    var rowsOnly = TH.rule !== 'all';
    var cellBorder = rowsOnly
      ? 'border:0;border-bottom:1px solid ' + LINE + ';'
      : 'border:1px solid ' + LINE + ';';
    return [
      '*{box-sizing:border-box;}',
      'html,body{margin:0;padding:0;background:#FFFFFF;color:' + INK + ';',
      "font-family:'Noto Sans JP','Hiragino Kaku Gothic ProN','Yu Gothic',sans-serif;",
      '-webkit-print-color-adjust:exact;print-color-adjust:exact;}',
      '.sheet{width:190mm;min-width:190mm;margin:0 auto;padding:12mm 10mm;position:relative;}',
      '.sheet + .sheet{border-top:1px dashed ' + LINE + ';}',
      '@media print{.sheet{page-break-after:always;break-after:page;border-top:0;}',
      '.sheet:last-child{page-break-after:auto;break-after:auto;}}',

      /* ★日付・No. の塊は題名より上（実物と同じ並び）。題名を少し下げて場所を空ける★ */
      '.ttl{font-size:22pt;letter-spacing:' + TH.titleSpacing + ';text-align:center;color:' + INK + ';',
      'margin:11mm 0 8mm;font-weight:700;}',

      /* 日付・No.（右上）。★ラベルの後ろは全角スペース＝字間はそれで作る★ */
      '.meta{position:absolute;top:12mm;right:10mm;font-size:9pt;color:' + SUB + ';text-align:right;}',
      '.meta-l{display:block;white-space:nowrap;line-height:1.9;}',

      /* 宛名（左）／自社（右）。★表の2列＝幅が足りなくても文が縦に割れない★
         ★下線は引かない（うちの紙は引いていない）★ */
      '.party{width:100%;border-collapse:collapse;margin:0 0 5mm;table-layout:fixed;}',
      '.party td{vertical-align:top;padding:0;}',
      '.party-to{width:56%;min-width:80mm;}',
      '.party-from{width:44%;min-width:60mm;text-align:right;}',
      '.to-name{font-size:15pt;font-weight:700;display:block;line-height:1.6;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.to-sub{font-size:9.5pt;color:' + SUB + ';line-height:1.8;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.from-name{font-size:11pt;font-weight:700;line-height:1.6;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.from-sub{font-size:9pt;color:' + SUB + ';line-height:1.6;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.seal{width:18mm;height:18mm;object-fit:contain;margin-top:2mm;}',
      '.pageno{font-size:9.5pt;color:' + SUB + ';margin:0 0 3mm;}',

      /* 挨拶。★block＋十分な幅＝1文字ずつ縦に割れない★ */
      '.lead{margin:0 0 5mm;}',
      '.lead-l{display:block;width:100%;min-width:80mm;font-size:9.5pt;color:' + SUB + ';',
      'line-height:1.9;white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      /* ★御請求金額＝枠なし。ラベル（小）＋金額（大）＋下に細い線★ */
      '.grand{margin:0 0 6mm;padding:0 0 2.2mm;border-bottom:1.2pt solid ' + ACCENT + ';',
      'display:block;width:100%;max-width:120mm;white-space:nowrap;}',
      '.grand-l{font-size:12pt;font-weight:700;color:' + INK + ';}',
      '.grand-v{font-size:20pt;font-weight:700;color:' + TH.grandInk + ';margin-left:12mm;',
      "font-family:'DM Mono',ui-monospace,monospace;letter-spacing:.02em;}",

      /* 表の上の小さなキャプション【…】 */
      '.cap{font-size:9pt;color:' + SUB + ';margin:0 0 2mm;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      /* 明細。★縦の罫は引かない・見出しは薄い地・表の下を1本で締める★ */
      '.items{width:100%;table-layout:fixed;border-collapse:collapse;font-size:9.5pt;margin:0 0 4mm;',
      'border-bottom:0.9pt solid ' + ACCENT + ';}',
      '.items th{background:' + TH.headBg + ';color:' + TH.headInk + ';font-weight:700;font-size:9pt;',
      'border:0;padding:2mm;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}',
      '.items td{' + cellBorder + 'padding:2mm;vertical-align:top;line-height:1.6;}',
      '.items .c-left{text-align:left;}',
      '.items .c-center{text-align:center;}',
      '.items .c-right{text-align:right;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.items .c-wrap{word-break:normal;overflow-wrap:break-word;}',
      '.items .c-nowrap{white-space:nowrap;}',
      '.c-memo{display:block;font-size:8.5pt;color:' + SUB + ';line-height:1.6;margin-top:.5mm;}',
      '.c-empty{text-align:center;color:' + SUB + ';padding:6mm 2mm;}',

      /* ★足元＝左に振込先／右に合計（表の2列＝文が縦に割れない）★ */
      '.foot{width:100%;border-collapse:collapse;table-layout:fixed;margin:0;}',
      '.foot td{vertical-align:top;padding:0;}',
      '.foot-l{width:52%;min-width:70mm;padding-right:6mm;}',
      '.foot-r{width:48%;min-width:70mm;}',

      /* ★小計/消費税/合計＝右下・枠なし・合計の上に線★ */
      '.sums{border-collapse:collapse;font-size:10pt;width:100%;margin:0 0 5mm auto;}',
      '.sums th{text-align:left;color:' + SUB + ';font-weight:400;border:0;',
      'padding:1.4mm 3mm;white-space:nowrap;}',
      '.sums td{text-align:right;border:0;padding:1.4mm 3mm;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.sums-g th{border-top:0.9pt solid ' + ACCENT + ';font-size:12pt;font-weight:700;color:' + INK + ';}',
      '.sums-g td{border-top:0.9pt solid ' + ACCENT + ';font-size:14pt;font-weight:700;color:' + TH.grandInk + ';}',

      /* （内訳）★枠で囲まない★ */
      '.bd{margin:0 0 5mm;}',
      '.bd-h{font-size:8.5pt;color:' + SUB + ';margin:0 0 1mm;}',
      '.rates{border-collapse:collapse;font-size:9pt;width:100%;}',
      '.rates th{color:' + SUB + ';border:0;border-bottom:1px solid ' + LINE + ';',
      'padding:1.2mm 3mm;white-space:nowrap;text-align:left;font-weight:400;}',
      '.rates td{border:0;border-bottom:1px solid ' + LINE + ';padding:1.2mm 3mm;text-align:right;',
      "white-space:nowrap;font-family:'DM Mono',ui-monospace,monospace;}",
      '.rates tbody th{color:' + INK + ';}',
      '.r-none{text-align:center;color:' + SUB + ';}',

      /* 振込先・備考。★箱で囲まない★（文の幅だけは確保する） */
      '.note{margin:0 0 3mm;}',
      '.note-h{font-size:8.5pt;color:' + SUB + ';margin-bottom:.8mm;}',
      '.note-b{display:block;width:100%;min-width:60mm;font-size:9.5pt;line-height:1.9;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      /* 次の紙へ送る時 */
      '.cont{margin:0 0 4mm;text-align:right;}',
      '.cont-l{display:block;font-size:9.5pt;color:' + SUB + ';white-space:nowrap;}',
      '.cont-v{margin-left:6mm;color:' + INK + ";font-family:'DM Mono',ui-monospace,monospace;}",
      '.cont-n{display:block;font-size:10pt;color:' + INK + ';margin-top:1.5mm;white-space:nowrap;}',

      '@page{size:A4;margin:0;}',
      '@media print{.sheet{width:210mm;min-width:210mm;padding:14mm 12mm;}}',
    ].join('');
  }

  return {
    build: build, css: css, esc: esc, yen: yen, comma: comma,
    dateStr: dateStr, jpDate: jpDate, honorOf: honorOf, taxLabel: taxLabel,
    paginate: paginate, TEMPLATE_ID: TEMPLATE_ID,
    ROWS_FIRST: ROWS_FIRST, ROWS_REST: ROWS_REST,
  };
});
