/* seikyu-paper.js — ★自作テンプレ std1（紙の中身）を作る唯一の場所★
 * ==============================================================================
 * 何をする物か:
 *   請求書1通ぶんの「紙」を、そのまま印刷できる1枚のHTMLとして返す。
 *   ★画面のプレビューも、印刷も、同じこの文字列を使う★＝見えていた物と刷れる物が食い違わない。
 *
 * ★印刷は「紙だけの新しい窓」で刷る（司さんの決まり）★
 *   画面に @media print を掛けて隠す作りは、隠し忘れが必ず出る・端末で結果が変わる。
 *   ここが返すHTMLには ★アプリの画面が1バイトも入っていない★ ので、その窓を刷れば紙だけが出る。
 *
 * ★注意書き・但し書きは「箱で潰さない」（過去2回踏んだ）★
 *   flex/grid の子に長い日本語を入れると、幅が足りない時に ★1文字ずつ縦に割れる★。
 *   だからこの紙では、文が入る所に flex/grid を使わない（表と block だけ）。
 *   さらに文の箱には min-width を持たせ、white-space を折り返し可にしている。
 *   機械での見張り = seikyu/tests/seikyu-paper.test.mjs
 *
 * ★税率の数字を1つも書かない★
 *   区分は totals.byRate（seikyu-tax.js が出した物）をそのまま並べる。
 *   率が変わってもこのファイルは直さない。
 *
 * ★画面に依らない（DOMを1つも触らない）／時計も乱数も持たない★＝素のNodeで全パターン回せる。
 *
 * 【利用】ブラウザ window.SeikyuPaper ／ Node require('./seikyu-paper.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SeikyuPaper = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var TEMPLATE_ID = 'std1';

  /* 色は★直hex★（このリポジトリの現行。:root 変数は飲み屋だけの話）。
     ★#1A4A2E は使わない（濃すぎる）。全アプリ #2E7D54★ */
  var INK = '#24422F';       // 本文
  var SUB = '#5D7C6B';       // 補足
  var LINE = '#C9DED2';      // 罫線
  var ACCENT = '#2E7D54';    // 見出し・強調
  var BAND = '#EEF7F1';      // 帯の下地

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* 円。★数にならない物を 0 にしない（取れなかったを 0 と作り分ける）★ */
  function yen(v) {
    var n = Number(v);
    if (!Number.isFinite(n)) return '—';
    var s = Math.abs(Math.trunc(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return (n < 0 ? '-' : '') + s;
  }
  /* 数量・単価は小数が入りうる（0.5時間 など）。整数なら桁区切りだけ付ける */
  function num(v) {
    if (v === undefined || v === null || v === '') return '';
    var n = Number(v);
    if (!Number.isFinite(n)) return esc(v);
    if (Number.isInteger(n)) return n.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return String(n);
  }
  /* 'YYYY-MM-DD' → '2026年9月30日'。読めなければ空（★勝手に今日を入れない★） */
  function jpDate(ymd) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
    if (!m) return '';
    return +m[1] + '年' + (+m[2]) + '月' + (+m[3]) + '日';
  }

  /* 敬称。★取引先マスタは既に hub が `keisho` で書いている★ ので両方を見る。
     「（なし）」は付けない（空文字にする）。 */
  function honorOf(p) {
    var h = (p && (p.honor || p.keisho)) || '';
    if (!h || h === '（なし）' || h === '(なし)' || h === 'なし') return '';
    return h;
  }

  /* ═══ 紙1枚 ═══
   * build({ inv, tax, partner, org, title })
   *   inv    … 請求書の行（no / issue_ymd / due_ymd / doc_type / lines / data.subject / data.memo …）
   *   tax    … seikyu-tax.compute の返り（byRate / subtotal / taxTotal / grandTotal / exempt）
   *   partner… 宛先（発行済みなら snapshot.partner・下書きならマスタの data）
   *   org    … 自社（発行済みなら snapshot.org・下書きなら pay_org の data）
   * 返り = { html, title, templateId }
   */
  function build(o) {
    o = o || {};
    var inv = o.inv || {};
    var tax = o.tax || {};
    var p = o.partner || {};
    var g = o.org || {};
    var isQuote = (inv.doc_type === 'quote');
    var heading = isQuote ? '見 積 書' : '請 求 書';
    var totalLabel = isQuote ? 'お見積金額' : 'ご請求金額';
    var docTitle = (o.title || (heading.replace(/ /g, '') + (inv.no ? ' ' + inv.no : '')));

    var lines = Array.isArray(tax.lines) ? tax.lines : [];
    var byRate = Array.isArray(tax.byRate) ? tax.byRate : [];
    var exemptBase = (tax.exempt && Number(tax.exempt.base)) || 0;
    var taxMode = inv.tax_mode === 'inclusive' ? '内税' : '外税';

    /* ── 明細行。★1行も無い時に空の表を出さない（何も無いと分かる文を出す）★ */
    var rowsHtml = lines.length
      ? lines.map(function (ln, i) {
        return '<tr>'
          + '<td class="c-no">' + (i + 1) + '</td>'
          + '<td class="c-name">' + esc(ln.name || '') + (ln.memo ? '<span class="c-memo">' + esc(ln.memo) + '</span>' : '') + '</td>'
          + '<td class="c-qty">' + num(ln.qty) + '</td>'
          + '<td class="c-unit">' + esc(ln.unit || '') + '</td>'
          + '<td class="c-price">' + (ln.price === undefined || ln.price === null || ln.price === '' ? '' : yen(ln.price)) + '</td>'
          + '<td class="c-amt">' + yen(ln.amount) + '</td>'
          + '<td class="c-rate">' + (Number(ln.rate) === 0 ? '—' : num(ln.rate) + '%') + '</td>'
          + '</tr>';
      }).join('')
      : '<tr><td class="c-empty" colspan="7">明細がまだ1行もありません</td></tr>';

    /* ── 税率ごとの区分（適格請求書の要件）。★率の数字はデータから出す★ */
    var rateRows = byRate.map(function (b) {
      return '<tr><th>' + num(b.pct) + '% 対象</th>'
        + '<td class="r-base">' + yen(b.base) + '</td>'
        + '<td class="r-tax">' + yen(b.tax) + '</td></tr>';
    }).join('');
    if (exemptBase !== 0) {
      rateRows += '<tr><th>消費税の対象外</th><td class="r-base">' + yen(exemptBase) + '</td><td class="r-tax">—</td></tr>';
    }
    if (!rateRows) rateRows = '<tr><td class="r-none" colspan="3">区分はまだありません</td></tr>';

    /* ── 自社の振込先・備考。★文は表と block にだけ入れる（flex/grid に入れない）★ */
    var bankBlock = g.bank
      ? '<div class="note-box"><div class="note-h">お振込先</div><div class="note-b">' + esc(g.bank).replace(/\n/g, '<br>') + '</div></div>'
      : '';
    var memo = (inv.data && inv.data.memo) || '';
    var memoBlock = memo
      ? '<div class="note-box"><div class="note-h">備考</div><div class="note-b">' + esc(memo).replace(/\n/g, '<br>') + '</div></div>'
      : '';

    var subject = (inv.data && inv.data.subject) || '';
    var due = jpDate(inv.due_ymd);

    var html = ''
      + '<!DOCTYPE html>\n<html lang="ja"><head><meta charset="UTF-8">'
      + '<meta name="viewport" content="width=device-width, initial-scale=1">'
      + '<title>' + esc(docTitle) + '</title>'
      + '<style>' + css() + '</style></head><body>'
      + '<div class="sheet">'

      /* 見出し */
      + '<h1 class="ttl">' + heading + '</h1>'
      + '<table class="meta"><tbody>'
      + '<tr><th>' + (isQuote ? '見積番号' : '請求番号') + '</th><td>' + (esc(inv.no) || '（未採番）') + '</td></tr>'
      + '<tr><th>' + (isQuote ? '見積日' : '請求日') + '</th><td>' + (jpDate(inv.issue_ymd) || '（未入力）') + '</td></tr>'
      + (due ? '<tr><th>お支払期限</th><td>' + due + '</td></tr>' : '')
      + '</tbody></table>'

      /* 宛先（左）と自社（右）。★2段組みは表で作る★
         flex/grid だと幅が足りない時に長い会社名が1文字ずつ縦に割れる（前科2回）。 */
      + '<table class="party"><tbody><tr>'
      + '<td class="party-to">'
      + '<div class="to-name">' + (esc(p.name) || '（取引先が未選択）') + (honorOf(p) ? '<span class="to-honor">' + esc(honorOf(p)) + '</span>' : '') + '</div>'
      + (p.person ? '<div class="to-sub">' + esc(p.person) + ' 様</div>' : '')
      + (p.zip ? '<div class="to-sub">〒' + esc(p.zip) + '</div>' : '')
      + (p.addr ? '<div class="to-sub">' + esc(p.addr) + '</div>' : '')
      + '</td>'
      + '<td class="party-from">'
      + '<div class="from-name">' + (esc(g.yago) || '（自社情報が未入力）') + '</div>'
      + (g.addr ? '<div class="from-sub">' + esc(g.addr) + '</div>' : '')
      + (g.tel ? '<div class="from-sub">TEL ' + esc(g.tel) + '</div>' : '')
      + (g.invoiceNo ? '<div class="from-sub">登録番号 ' + esc(g.invoiceNo) + '</div>' : '')
      + (g.sealDataUrl ? '<img class="seal" src="' + esc(g.sealDataUrl) + '" alt="">' : '')
      + '</td>'
      + '</tr></tbody></table>'

      /* 合計（いちばん大きく） */
      + '<div class="grand">'
      + '<div class="grand-l">' + totalLabel + '</div>'
      + '<div class="grand-v">' + yen(tax.grandTotal) + '<span class="grand-en">円</span></div>'
      + '<div class="grand-n">（' + taxMode + '／消費税込み）</div>'
      + '</div>'

      + (subject ? '<div class="subject"><span class="subject-h">件名</span><span class="subject-b">' + esc(subject) + '</span></div>' : '')

      /* 明細 */
      + '<table class="items"><thead><tr>'
      + '<th class="c-no">#</th><th class="c-name">品名・内容</th><th class="c-qty">数量</th>'
      + '<th class="c-unit">単位</th><th class="c-price">単価</th><th class="c-amt">金額</th><th class="c-rate">税率</th>'
      + '</tr></thead><tbody>' + rowsHtml + '</tbody></table>'

      /* 合計欄 ＋ 税率ごとの区分 */
      + '<table class="sums"><tbody>'
      + '<tr><th>小計</th><td>' + yen(tax.subtotal) + '</td></tr>'
      + '<tr><th>消費税</th><td>' + yen(tax.taxTotal) + '</td></tr>'
      + '<tr class="sums-g"><th>合計</th><td>' + yen(tax.grandTotal) + '</td></tr>'
      + '</tbody></table>'
      + '<table class="rates"><thead><tr><th>区分</th><th>対象額</th><th>消費税</th></tr></thead>'
      + '<tbody>' + rateRows + '</tbody></table>'

      + bankBlock
      + memoBlock
      + '</div></body></html>';

    return { html: html, title: docTitle, templateId: TEMPLATE_ID };
  }

  /* ── 紙の見た目 ──────────────────────────────────────────────
     ★文が入る所に flex/grid を使わない★（1文字ずつ縦に割れる事故を作らない）。
     幅は mm で決める＝画面の幅に引きずられない。 */
  function css() {
    return [
      '*{box-sizing:border-box;}',
      'html,body{margin:0;padding:0;background:#FFFFFF;color:' + INK + ';',
      "font-family:'Noto Sans JP','Hiragino Kaku Gothic ProN','Yu Gothic',sans-serif;",
      '-webkit-print-color-adjust:exact;print-color-adjust:exact;}',
      '.sheet{width:190mm;min-width:190mm;margin:0 auto;padding:12mm 10mm;position:relative;}',

      '.ttl{font-size:22pt;letter-spacing:.32em;text-align:center;color:' + ACCENT + ';',
      'margin:0 0 8mm;font-weight:700;}',

      /* 番号・日付（右上）。表なので文が縦に割れない */
      '.meta{position:absolute;top:12mm;right:10mm;border-collapse:collapse;font-size:9pt;}',
      '.meta th{color:' + SUB + ';font-weight:400;text-align:left;padding:1px 8px 1px 0;white-space:nowrap;}',
      '.meta td{text-align:right;padding:1px 0;white-space:nowrap;}',

      /* 宛先（左）／自社（右）。★表の2列＝幅が足りなくても文が縦に割れない★ */
      '.party{width:100%;border-collapse:collapse;margin:0 0 6mm;table-layout:fixed;}',
      '.party td{vertical-align:top;padding:0;}',
      '.party-to{width:56%;min-width:80mm;}',
      '.party-from{width:44%;min-width:60mm;text-align:right;}',
      '.to-name{font-size:14pt;font-weight:700;border-bottom:1px solid ' + INK + ';',
      'display:block;padding:0 0 2mm;line-height:1.5;word-break:normal;overflow-wrap:break-word;}',
      '.to-honor{font-size:11pt;font-weight:400;margin-left:6px;}',
      '.to-sub{font-size:9.5pt;color:' + SUB + ';line-height:1.7;margin-top:1mm;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.from-name{font-size:11pt;font-weight:700;line-height:1.6;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.from-sub{font-size:9pt;color:' + SUB + ';line-height:1.6;',
      'word-break:normal;overflow-wrap:break-word;}',
      '.seal{width:18mm;height:18mm;object-fit:contain;margin-top:2mm;}',

      /* 合計（帯） */
      '.grand{background:' + BAND + ';border:1px solid ' + LINE + ';border-radius:2mm;',
      'padding:4mm 6mm;margin:0 0 5mm;text-align:center;}',
      '.grand-l{font-size:9.5pt;color:' + SUB + ';letter-spacing:.08em;}',
      '.grand-v{font-size:24pt;font-weight:700;color:' + ACCENT + ';line-height:1.25;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.grand-en{font-size:12pt;margin-left:2mm;}',
      /* ★この1行が「1文字ずつ縦」になった前科の形。block ＋ 折返し可 ＋ 十分な幅で潰さない */
      '.grand-n{display:block;width:100%;min-width:60mm;font-size:8.5pt;color:' + SUB + ';',
      'line-height:1.7;white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      /* 件名 */
      '.subject{margin:0 0 4mm;font-size:10pt;line-height:1.8;}',
      '.subject-h{display:inline-block;min-width:14mm;color:' + SUB + ';font-size:9pt;}',
      '.subject-b{font-weight:700;}',

      /* 明細 */
      '.items{width:100%;border-collapse:collapse;font-size:9.5pt;margin:0 0 4mm;}',
      '.items th{background:' + BAND + ';color:' + ACCENT + ';font-weight:700;font-size:9pt;',
      'border:1px solid ' + LINE + ';padding:2mm 2mm;white-space:nowrap;}',
      '.items td{border:1px solid ' + LINE + ';padding:2mm;vertical-align:top;line-height:1.6;}',
      '.items .c-no{width:8mm;text-align:center;color:' + SUB + ';}',
      '.items .c-name{width:auto;min-width:50mm;word-break:normal;overflow-wrap:break-word;}',
      '.items .c-qty,.items .c-price,.items .c-amt{text-align:right;white-space:nowrap;}',
      '.items .c-qty{width:16mm;}.items .c-unit{width:12mm;text-align:center;}',
      '.items .c-price{width:24mm;}.items .c-amt{width:26mm;}',
      '.items .c-rate{width:14mm;text-align:center;white-space:nowrap;}',
      '.c-memo{display:block;font-size:8.5pt;color:' + SUB + ';line-height:1.6;margin-top:.5mm;}',
      '.c-empty{text-align:center;color:' + SUB + ';padding:6mm 2mm;}',

      /* 合計欄（右）と区分（左）。★float も flex も使わず、幅つきの表を2つ並べる★ */
      '.sums{border-collapse:collapse;font-size:10pt;width:70mm;margin:0 0 4mm auto;}',
      '.sums th{text-align:left;color:' + SUB + ';font-weight:400;border:1px solid ' + LINE + ';',
      'padding:1.6mm 3mm;background:#FFFFFF;white-space:nowrap;}',
      '.sums td{text-align:right;border:1px solid ' + LINE + ';padding:1.6mm 3mm;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.sums-g th,.sums-g td{background:' + BAND + ';font-weight:700;color:' + ACCENT + ';font-size:11pt;}',

      '.rates{border-collapse:collapse;font-size:9pt;width:80mm;margin:0 0 5mm;}',
      '.rates th{background:' + BAND + ';color:' + ACCENT + ';border:1px solid ' + LINE + ';',
      'padding:1.4mm 3mm;white-space:nowrap;text-align:left;font-weight:700;}',
      '.rates td{border:1px solid ' + LINE + ';padding:1.4mm 3mm;text-align:right;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.rates tbody th{background:#FFFFFF;font-weight:400;color:' + INK + ';}',
      '.r-none{text-align:center;color:' + SUB + ';}',

      /* 注意書き・振込先。★ここが縦に割れた前科の形。block＋min-width で潰さない★ */
      '.note-box{border:1px solid ' + LINE + ';border-radius:2mm;padding:3mm 4mm;margin:0 0 3mm;',
      'width:100%;min-width:80mm;}',
      '.note-h{font-size:8.5pt;color:' + ACCENT + ';font-weight:700;margin-bottom:1mm;}',
      '.note-b{display:block;width:100%;min-width:60mm;font-size:9.5pt;line-height:1.9;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      /* 印刷。A4・余白は紙側に任せ、こちらは色を落とさない */
      '@page{size:A4;margin:0;}',
      '@media print{.sheet{width:210mm;min-width:210mm;padding:14mm 12mm;}}',
    ].join('');
  }

  return { build: build, css: css, esc: esc, yen: yen, jpDate: jpDate, honorOf: honorOf, TEMPLATE_ID: TEMPLATE_ID };
});
