/* seikyu-paper.js — ★請求書の紙を刷る唯一の場所★
 * ==============================================================================
 * ★焼き付けてよい型は「法律」だけ★（2026-08-10 指示役の方向の訂正）
 *   国税庁 適格請求書の記載事項6つ ＝ 業種に関係なく必ず要る＝ここは動かさない。
 *     ①発行する側の名称と登録番号 ②取引年月日 ③取引の内容（軽減税率の対象である旨）
 *     ④税率ごとの対価の額と適用税率 ⑤税率ごとの消費税額 ⑥受け取る側の名称
 *   ★これ以外の言い回し（挨拶文・ラベルの字間・和暦・¥・ページ送りの言葉）は
 *     「出せる」ようにしてあるだけで、「出さなければ赤」にはしていません。★
 *   理由: 代行請求アプリは ★源泉なし・繰越なし・非課税なし・1税率★＝一番 単純な1業種の紙。
 *         見本としては良いが、そこに錠を掛けると 士業（源泉）・掛け売り（繰越）・
 *         不動産（非課税）の客が全部 落ちる。★見本は写す相手ではない。★
 *
 * ★見た目の決めごと（うちの家の作り。理由は「他所がそうだから」ではない）★
 *   ・金額は ★枠で囲まず、大きく出して下に細い線★
 *     ＝紙の主役は金額。塗った箱に入れると「囲みの中の一情報」に見えて主役が下がる。
 *   ・お振込先・備考も ★囲まない★。囲みが増えるほど、どこを読めばよいか分からなくなる。
 *   ・明細は ★縦の罫を引かず、横の細い罫だけ★。列を会社が足せるので、
 *     縦罫があると列を増やすたびに紙が檻のようになる。
 *   ・色は Exally／Kyually と同じ緑（css/exally-ui.css と同じ13色から）。
 *     ※代行請求（ダイコメの製品）は今 事務所の青(#007AFF)ですが、それは別の家の色です。
 *
 * ★税率の数字を1つも書かない★（区分は seikyu-tax.js が出した物を並べるだけ）
 * ★画面に依らない（DOMを1つも触らない）／時計も乱数も持たない★
 *
 * 見本として読んだ物（写してはいない）: daikou-seikyu-test/invoice-pdf.js
 *
 * 【利用】ブラウザ window.SeikyuPaper ／ Node require('./seikyu-paper.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./seikyu-cols.js'), require('./seikyu-carry.js'), require('./seikyu-doc.js'));
  else root.SeikyuPaper = factory(root.SeikyuCols, root.SeikyuCarry, root.SeikyuDoc);
})(typeof self !== 'undefined' ? self : this, function (COLS, CARRY, DOC) {
  'use strict';
  if (!COLS) throw new Error('seikyu-cols.js を先に読んでください');
  if (!CARRY) throw new Error('seikyu-carry.js を先に読んでください');
  if (!DOC) throw new Error('seikyu-doc.js を先に読んでください');

  /* ★繰越の行の名前は「繰越 lib」が持ち主★
     紙の側に書き写すと、片方だけ直した時に画面と紙で名前が食い違う。 */
  var CARRY_ROWS = CARRY.ROWS;

  var TEMPLATE_ID = 'std1';

  /* 1枚に載る明細の行数。★ここを超えたら次の紙へ送る（黙って切らない）★
     1枚目は 宛名・挨拶・御請求金額 が乗るぶん少ない。 */
  /* ★項目が少なくても 紙の顔を同じにする★
     うちの実物が すでにそうなっている（実測 2026-08-15）:
       黒田空調 … 入っているのは6行／枠は ★12〜41行＝30行★
       ENEOS   … 入っているのは2行／枠は ★同じ30行★
       八木工業 … 控除の枠は ★E17:H20＝4行★（1行しか使っていない）
     ＝★中身が2行でも6行でも、毎月 同じ場所に同じ物が来る★。
       経理の人は「いつも同じ場所」を見るので、動くと毎回 探し直しになる。
     ★足りない行は空の枠のまま残す（罫線は消さない）★
     ★固定にするのは「紙」だけ★＝入力の画面は今までどおり可変（空欄を並べて埋めさせない）。 */
  /* ★行の高さは1か所で決める★＝左（明細）と右（控除）の罫線がずれない
     （2026-08-15 司さんの指摘：右の空の枠の横線が左と合っていなかった） */
  var ROW_H = '7.2mm';
  var PAPER_ROWS = 30;        // 明細の枠（1ページぶん）★会社が変えられる★
  var DEDUCT_ROWS = 4;        // 差し引く（控除）の枠 ★会社が変えられる★
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
    grandGo: 'ご',        // 「ご請求金額（税込）」（様式で「御」にもできる）
  };
  function themeOf(t) { return Object.assign({}, THEME, t || {}); }

  var DEFAULT_COLS = {
    items: ['#', '品名・内容', '数量', '単位', '単価', '金額', '消費税'],
    widths: { '#': 28, '品名・内容': 220, '数量': 56, '単位': 44, '単価': 80, '金額': 100, '消費税': 72 },
    aligns: {},
  };

  function esc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  /* ★人が読む文として使えるか★
     String() は 物 を渡されると "[object Object]" を作る。それを紙に刷ると、
     客の手元に [object Object] と書かれた請求書が届く
     （2026-08-11 実機で発生：お振込先に 物 が入っていた）。
     ★読めない物は紙に出さない★（ここは法定の項目ではなく、書いてあれば読む物）。 */
  function textOf(v) {
    if (v == null) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return Number.isFinite(v) ? String(v) : '';
    return '';
  }

  /* ★紙の金額は ¥ 記号（画面は「1,100 円」。二重に付けない）★
     数にならない物は 0 にしない（取れなかったを 0 と作り分ける）。 */
  function yen(v) {
    var n = Number(v);
    if (!Number.isFinite(n)) return '—';
    return '¥' + Math.round(n).toLocaleString('ja-JP');
  }
  /* 表の中の数（¥ を付けない・桁区切りだけ。桁が詰まって読みにくくなるため） */
  function comma(v) {
    if (v === undefined || v === null || v === '') return '';
    var n = Number(v);
    if (!Number.isFinite(n)) return esc(v);
    return n.toLocaleString('ja-JP');
  }
  function num(v) { return comma(v); }

  /* ★日付＝和暦か西暦を選べる（既定は西暦）★
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

  /* 角印の大きさ（mm）。10〜40に収める＝紙からはみ出す印を作らない */
  function sealMm(v) {
    var n = Number(v);
    if (!Number.isFinite(n)) return 21;
    return Math.max(10, Math.min(40, Math.round(n)));
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

    /* ★紙の種類は3つ★ 請求書／見積書／★領収書★
       領収書は doc_type ではない（棚を増やさない）＝入金1行から出す紙なので、
       ★呼ぶ側が docKind:'receipt' と receipt:{…} を渡す★。 */
    var rc = o.receipt || null;
    var kind = o.docKind || (inv.doc_type === 'quote' ? 'quote' : 'invoice');
    if (kind === 'receipt' && !rc) kind = 'invoice';        // 中身が無いのに領収書の顔をしない
    var isQuote = (kind === 'quote');
    var isReceipt = (kind === 'receipt');
    var heading = isReceipt ? '領　収　書' : isQuote ? '見　積　書' : '請　求　書';
    /* 金額のラベルは様式が持つ（ご／御）。★どちらでもよい＝縛らない★ */
    var go = TH.grandGo || 'ご';
    var grandLabel = isReceipt ? '領収金額（税込）' : go + (isQuote ? '見積金額（税込）' : '請求金額（税込）');
    var noLabel = 'No.　';
    var headNo = isReceipt ? String((rc && rc.no) || '') : (inv.no || '');
    var docTitle = (o.title || (heading.replace(/　/g, '') + (headNo ? ' ' + headNo : '')));

    var lines = Array.isArray(tax.lines) ? tax.lines : [];
    var byRate = Array.isArray(tax.byRate) ? tax.byRate : [];
    var exemptBase = (tax.exempt && Number(tax.exempt.base)) || 0;
    var nontaxBase = (tax.nontaxable && Number(tax.nontaxable.base)) || 0;
    /* ★控除（明細の外・税込の合計から引く）★
       ★ここを紙に出し忘れると、画面は 281,260 なのに 紙は 292,600 と書く★
       ＝★請求している額と、紙に書いた額が食い違う★（2026-08-15 スクショで実際に見つけた）。
       決め方は seikyu-doc.js が唯一の正（billedOf）。 */
    var deduct = (o.deduct === undefined || o.deduct === null) ? null : Number(o.deduct);
    var deductLines = Array.isArray(o.deductLines) ? o.deductLines : [];
    /* ★枠の本数（固定）★ 会社が変えられる。0以下は「詰める」＝昔の形に戻す道 */
    var frameRows = Math.max(0, Math.trunc(Number(
      (o.paperRows !== undefined ? o.paperRows : (inv.data && inv.data.paperRows))) || PAPER_ROWS));
    var dedRows = Math.max(0, Math.trunc(Number(
      (o.deductRows !== undefined ? o.deductRows : (inv.data && inv.data.deductRows))) || DEDUCT_ROWS));
    /* ★控除の話を渡されたか★
       渡されていない紙（見積・領収・古い呼び方）に「控除計（未確認）」を出さない。
       ★渡された時は 0件でも枠を出す★＝毎月 同じ顔（実物 八木は4行の枠に1行しか使っていない）。
       ★「出さない」も会社が選べる★（控除を使わない会社が多いため）。 */
    var deductGiven = (o.deduct !== undefined) || deductLines.length > 0 || o.showDeduct !== undefined
      || !!(inv.data && inv.data.deductions && inv.data.deductions.length);
    /* ★1カラム／2カラム★（給料明細と同じ選び方・同じ言い方）
         2カラム … 明細｜差し引く を ★横に並べる★（col2）
         1カラム … 明細 ↓ 差し引く を ★縦に積む★（col1・明細が多い／控除を使わない会社向け） */
    var layout = String((o.layout !== undefined ? o.layout : (inv.data && inv.data.paperLayout)) || 'col2');
    if (layout !== 'col1' && layout !== 'col2') layout = 'col2';
    var showDeduct = (o.showDeduct !== undefined) ? !!o.showDeduct
      : ((inv.data && inv.data.showDeduct) !== undefined ? !!inv.data.showDeduct : deductGiven);
    var gen = o.gensen || null;     // ★源泉徴収（引く紙だけ）★
    var carry = o.carry || null;    // ★繰越（前回の残り）★
    /* ★1ページに入るのは 枠の本数まで★（入り切らない時だけ次のページ） */
    var pages = paginate(lines, frameRows
      ? { rowsFirst: frameRows, rowsRest: frameRows }
      : o.page);
    var multi = pages.length > 1;

    /* 明細の見出し（items のとおり・その順） */
    var headHtml = spec.items.map(function (k, c) {
      return '<th class="c-col" style="width:' + colW[c].toFixed(4) + '%;text-align:' + COLS.alignOf(spec, k) + '">' + esc(k) + '</th>';
    }).join('');

    /* ★足りない行は「空の枠」のまま残す★（罫線を消さない・詰めない）
       ここで詰めると ★中身の本数で紙の顔が毎月 変わる★＝経理が毎回 探し直す。 */
    function blankRowsHtml(n) {
      if (n <= 0) return '';
      var tds = spec.items.map(function (k) {
        return '<td class="c-col c-' + COLS.alignOf(spec, k) + ' c-blank">&nbsp;</td>';
      }).join('');
      var out = '';
      for (var i = 0; i < n; i++) out += '<tr class="r-blank">' + tds + '</tr>';
      return out;
    }

    function rowsHtmlOf(pageLines, offset) {
      /* ★1行も無い時は「無い」と言う★（空の枠だけだと、消えたのか元から無いのか分からない）。
         言った上で ★残りは空の枠★＝紙の顔は同じ高さのまま。 */
      if (!pageLines.length) {
        return '<tr><td class="c-empty" colspan="' + spec.items.length + '">明細がまだ1行もありません</td></tr>'
          + blankRowsHtml(Math.max(0, frameRows - 1));
      }
      return pageLines.map(function (ln, i) {
        return '<tr>' + spec.items.map(function (k) {
          var cell = COLS.cellOf(ln, k, offset + i, spec);
          var al = COLS.alignOf(spec, k);
          var role = COLS.roleOfIn(spec, k);
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
      }).join('') + blankRowsHtml(frameRows - pageLines.length);
    }

    /* ── 頭（宛名・自社・日付） ★ラベルの後ろは全角スペース★ ── */
    function headBlock(pageIdx) {
      /* 並びは 請求日 → No. → お支払期限。
         ★番号は空でも欄を出す（「（未採番）」と書く）＝取れなかったを空欄にしない★ */
      /* 領収書は ★入金日★ が日付・★領収番号（請求番号＋枝番）★ が番号 */
      var ds = dateStr(isReceipt ? (rc && rc.ymd) : inv.issue_ymd, era);
      var dLabel = isReceipt ? '領収日　' : isQuote ? '見積日　' : '請求日　';
      var meta = '<div class="meta-l">' + dLabel + (ds || '（未入力）') + '</div>';
      meta += '<div class="meta-l">' + esc(noLabel) + (esc(headNo) || '（未採番）') + '</div>';
      // ★もう受け取った紙に「お支払期限」を出さない★
      if (!isReceipt && inv.due_ymd) meta += '<div class="meta-l">お支払期限　' + dateStr(inv.due_ymd, era) + '</div>';

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
        + (g.sealDataUrl ? '<img class="seal" style="width:' + sealMm(g.sealSizeMm) + 'mm;height:' + sealMm(g.sealSizeMm) + 'mm" src="' + esc(g.sealDataUrl) + '" alt="会社の印">' : '')
        + '</td></tr></tbody></table>'
        + (multi ? '<div class="pageno">' + (pageIdx + 1) + 'ページ目</div>' : '');
    }

    /* ── 挨拶（★下記の通り御請求申し上げます。★） ── */
    function leadBlock() {
      /* ★「◯年◯月分」＝請求日の★前月★★
         うちの実物32枚の標準様式は全部 =TEXT(EDATE(請求日,-1),"yyyy年m月分")。
         ここが「当月」だと、7月に出す紙に「7月分」と書いてしまう（★実物は6月分★）。
         決め方は seikyu-doc.js が唯一の正。読めない日付なら出さない（でっち上げない）。 */
      var lead = (inv.data && inv.data.lead) || '';
      if (!lead) lead = DOC.periodLabelOf(inv.issue_ymd);
      var greet = isQuote ? '下記の通り御見積申し上げます。' : '下記の通り御請求申し上げます。';
      return '<div class="lead">'
        + (lead ? '<div class="lead-l">' + esc(lead) + '</div>' : '')
        + '<div class="lead-l">' + greet + '</div>'
        + '</div>';
    }

    /* ── 御請求金額（★枠なし・ラベル＋大きい金額・下に線★） ── */
    function grandBlock() {
      /* ★領収書は「実際に受け取った額」★（請求額ではない）。
         一部だけ受け取った時に請求額を書くと、★受け取っていない金額の領収書★になる。 */
      if (isReceipt) {
        var got = Number(rc && rc.amount);
        return '<div class="grand">'
          + '<span class="grand-l">' + grandLabel + '</span>'
          + '<span class="grand-v">' + (Number.isFinite(got) ? yen(got) : '（未確認）') + '</span>'
          + '</div>';
      }
      /* ★見出しの額は「実際に請求している額」★＝繰越があれば足し、★控除があれば引いたあと★。
         ここだけ今回分のままにすると、下の 合計請求額／請求額 と食い違う。 */
      /* ★控除の話が在るのに読めない時は 頭の金額も数字にしない★
         （0として計算した額を大きく出すと ★引き忘れた紙★ になる） */
      var billed = (showDeduct && deduct === null) ? null : DOC.billedOf(tax, carry, deduct);
      return '<div class="grand">'
        + '<span class="grand-l">' + grandLabel + '</span>'
        + '<span class="grand-v">' + (billed === null ? '（未確認）' : yen(billed)) + '</span>'
        + '</div>';
    }

    /* ── 小計・消費税・合計（★枠なし・合計の上に線★）
         源泉徴収を引く紙は、合計の下に ★源泉徴収税額★ と ★差引お支払額★ を足す。
         ★足した2行が迷子にならないよう、合計の直下に続けて並べ、
           いちばん下（実際に払う額）を太くする★ */
    function totalsBlock() {
      /* ★締め★ 小計はブロックの合計が持っているので、ここは 消費税 → 合計 → 請求額。
         ★大きい数字は紙の頭に1つだけ★（給料明細の差引支給額と同じ）＝ここは全部 小さく。 */
      var rows = ''
        + '<tr><th>' + taxLabel(tax, inv.tax_mode) + '</th><td>' + yen(tax.taxTotal) + '</td></tr>'
        + '<tr class="sums-mid"><th>合計</th><td>' + yen(tax.grandTotal) + '</td></tr>';
      /* ★控除の「1行ずつ」は ②の枠が持つ★（同じ物を2か所に出さない）。
         ここには ★控除計 と 請求額★ だけを出す。★税額は動かさない★。
         ★控除が読めない時は 請求額も数字にしない★
           ＝0として計算した額を「請求額」と大きく出すと、★引き忘れた紙★ になる。 */
      if (showDeduct) {
        /* ★控除計は②のブロックが持つ★（同じ数を2回 足し算として出さない）。
           ここは ★引いたあとの「請求額」★ だけ＝締めの役目。 */
        var billedNet = (deduct === null) ? null : (tax.grandTotal - deduct);
        rows += '<tr class="sums-net"><th>請求額</th><td>'
          + (billedNet === null ? '（未確認）' : yen(billedNet)) + '</td></tr>';
      }
      if (gen && gen.on) {
        /* ★差引お支払額は「合計請求額（繰越こみ・控除ずみ）− 源泉」★
           gen.net は この1通だけで出した額なので、繰越があると足りない。
           順番は seikyu-doc.js が唯一の正。 */
        var pay = DOC.payableOf(tax, carry, gen, deduct);
        rows += '<tr class="sums-minus"><th>' + esc(gen.label) + '</th><td>-' + yen(gen.amount) + '</td></tr>'
          + '<tr class="sums-net"><th>' + esc(gen.netLabel) + '</th><td>' + (pay === null ? '（未確認）' : yen(pay)) + '</td></tr>';
      }
      return '<table class="sums"><tbody>' + rows + '</tbody></table>';
    }

    /* ── 繰越（前回の残り）★紙の頭・箱で囲まない★
         ★取れなかったを0にしない★＝読めなければ「入金は未確認」、初回は「前回の請求はありません」 */
    function carryBlock() {
      if (!carry) return '';
      /* ★「前回が無い」と「前回の入金が読めていない」は別物★
         初回に 前回請求額 — ／ 入金額 — ／ 繰越額 — の表を出すと、
         読めなかったのか元から無いのか、受け取った人に分からない。
         初回は ★1行だけ言う★（表は出さない）。 */
      if (carry.state === 'first') {
        return '<div class="carry"><div class="carry-n">' + esc(carry.label || '前回の請求はありません') + '</div></div>';
      }
      var v = function (x) { return (x === null || x === undefined) ? '（未確認）' : yen(x); };
      var rows = CARRY_ROWS.map(function (r) {
        return '<tr><th>' + r.label + '</th><td>' + v(carry[r.key]) + '</td></tr>';
      }).join('');
      var note = carry.label
        ? '<div class="carry-n">' + esc(carry.label) + (carry.prevNo ? '（前回 No.　' + esc(carry.prevNo) + '）' : '') + '</div>'
        : (carry.prevNo ? '<div class="carry-n">前回 No.　' + esc(carry.prevNo) + '</div>' : '');
      return '<div class="carry"><table class="carry-t"><tbody>' + rows + '</tbody></table>' + note + '</div>';
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

    /* ── ★② 差し引く（控除）★ ────────────────────────────────
       ★①の明細と はっきり別の枠★（見出し・罫線で分ける）。
       ★給料明細の「支給／控除／差引支給額」と同じ並び★＝Rakually の中で紙の作法をそろえる。
       ★税の外で引く★（①の中のマイナス行＝値引きは税の中なので、こちらに混ぜない）。
       ★中身が1行でも枠は固定★（実物 八木工業は E17:H20＝4行のうち1行しか使っていない）。 */
    /* ★給料明細と同じ作法★（kyuyo/js/render.js を読んで合わせた・2026-08-15）
         ・「支 給」「控 除」の ★見出しの下に線★（.st）＝左右の1行目が同じ高さから始まる
         ・足りない行は ★高さの決まった空行★ で埋める（rowsHTML の minCells と同じ考え方）
         ・★各ブロックの下にそのブロックの合計★（支給合計／控除合計＝上に線を引いた1行）
         ・★大きい数字は紙の頭に1つだけ★（差引支給額）＝ブロックの合計は小さく
       請求書では 支給→★ご請求の内訳★／控除→★差し引く★ に読み替える。 */
    function blockHead(t) { return '<div class="st">' + esc(t) + '</div>'; }
    function blockSum(label, value) {
      return '<table class="bsum"><tbody><tr><th>' + esc(label) + '</th><td>' + value + '</td></tr></tbody></table>';
    }

    function deductBlock() {
      if (!showDeduct) return '';
      var used = deductLines.length;
      var rows = deductLines.map(function (d) {
        var dv = Number(d && d.amount);
        return '<tr><th>' + (esc(d && d.name) || '控除') + '</th><td>'
          + (Number.isFinite(dv) ? '-' + yen(dv) : '（未確認）') + '</td></tr>';
      }).join('');
      var blanks = Math.max(0, dedRows - used);
      for (var i = 0; i < blanks; i++) rows += '<tr class="r-blank"><th>&nbsp;</th><td>&nbsp;</td></tr>';
      /* ★このブロックの合計はこのブロックが持つ★（給料明細の「控除合計」と同じ）。
         締めの「請求額」とは役目が違う（ブロックの足し算／払う額）。 */
      return '<div class="blk blk-ded">' + blockHead('差し引く')
        + '<table class="ded"><tbody>'
        + '<tr class="ded-hd"><th>内容</th><td>金額</td></tr>'
        + rows + '</tbody></table>'
        + blockSum('控除計', (deduct === null) ? '（未確認）' : (deduct ? '-' + yen(deduct) : yen(0)))
        + '</div>';
    }

    /* ── 足元（★左＝お振込先／右＝小計・消費税・合計＋（内訳）★）
         左右に並べるのは、紙の下半分を1列で長くしないため。
         ★2段組みは表で作る★（flex だと文が1文字ずつ縦に割れる） */
    function footerBlock() {
      var left = '';
      var bank = textOf(g.bank);
      if (bank) left += '<div class="note"><div class="note-h">お振込先</div><div class="note-b">' + esc(bank).replace(/\n/g, '<br>') + '</div></div>';
      var memo = textOf(inv.data && inv.data.memo);
      if (memo) left += '<div class="note"><div class="note-h">備考</div><div class="note-b">' + esc(memo).replace(/\n/g, '<br>') + '</div></div>';
      return '<table class="foot"><tbody><tr>'
        + '<td class="foot-l">' + left + '</td>'
        + '<td class="foot-r">' + breakdownBlock() + '</td>'
        + '</tr></tbody></table>';
    }

    var subject = (inv.data && inv.data.subject) || '';
    var caption = subject || (inv.data && inv.data.tableTitle) || '';

    /* ── 領収書の中身 ────────────────────────────────────────────
       ★明細も内訳も出さない★
         受け取ったのは「その1回ぶんのお金」で、明細は請求書の側にある。
         一部だけ受け取った紙に税率ごとの区分を出すと ★按分＝嘘の数字★ になる。
         だから ★どの請求の代金か（但し書き）で紐づける★。
       ★消費税額を区分して書くのは、はっきりしている時だけ★（国税庁 No.7124）
         書けた時は ★印紙の判定でも その分を記載金額から外す★（呼ぶ側が taxSeparate を立てる）。 */
    function receiptBody() {
      var but = textOf(rc.note) || (inv.no ? ('請求書 ' + inv.no + ' の代金として') : '上記代金として');
      var sep = !!rc.taxSeparate && Number(rc.taxTotal) > 0;
      var rows = '';
      if (sep) {
        var base = Number(rc.amount) - Number(rc.taxTotal);
        rows = '<table class="sums"><tbody>'
          + '<tr><th>税抜金額</th><td>' + yen(base) + '</td></tr>'
          + '<tr><th>消費税額等</th><td>' + yen(Number(rc.taxTotal)) + '</td></tr>'
          + '<tr class="sums-g"><th>合計</th><td>' + yen(Number(rc.amount)) + '</td></tr>'
          + '</tbody></table>';
      }
      /* ★印紙の注意は「要る時だけ」出す★。判定は seikyu-doc.js が唯一の正
         （区分して書けた時は、その消費税額を記載金額から外して数える＝No.7124）。 */
      var stamp = DOC.stampNote({ amount: Number(rc.amount), taxTotal: Number(rc.taxTotal), taxSeparate: sep });
      var memo = textOf(inv.data && inv.data.memo);
      return '<div class="rc-but"><span class="rc-but-l">但し　</span>'
        + '<span class="rc-but-b">' + esc(but) + '</span></div>'
        + '<div class="rc-ack">上記正に領収いたしました。</div>'
        + (rows ? '<div class="rc-sums">' + rows + '</div>' : '')
        + (rc.method ? '<div class="rc-way">お支払方法　' + esc(rc.method) + '</div>' : '')
        + (memo ? '<div class="note"><div class="note-h">備考</div><div class="note-b">' + esc(memo).replace(/\n/g, '<br>') + '</div></div>' : '')
        + (stamp ? '<div class="rc-stamp">' + esc(stamp) + '</div>' : '');
    }

    if (isReceipt) {
      var rcHtml = '<div class="sheet">' + headBlock(0) + grandBlock() + receiptBody() + '</div>';
      return {
        html: '<!DOCTYPE html>\n<html lang="ja"><head><meta charset="UTF-8">'
          + '<meta name="viewport" content="width=device-width, initial-scale=1">'
          + '<title>' + esc(docTitle) + '</title>'
          + '<style>' + css(TH) + '</style></head><body>' + rcHtml + '</body></html>',
        title: docTitle, templateId: inv.template_id || TEMPLATE_ID,
        cols: spec, colWidths: colW, pages: 1, docKind: 'receipt',
      };
    }

    /* ── 紙を組む ── */
    var sheets = pages.map(function (pageLines, idx) {
      var last = (idx === pages.length - 1);
      var offset = pages.slice(0, idx).reduce(function (a, x) { return a + x.length; }, 0);
      var pageSum = pageLines.reduce(function (a, x) { return a + (Number(x.amount) || 0); }, 0);
      var body = ''
        + headBlock(idx)
        /* ★繰越は1ページ目の金額のすぐ下★（前回の残りを含む額なので、金額の根拠として先に見せる）
           ここに並べないと carryBlock() は作られるだけで紙に載らない
           （2026-08-11：関数はあるのに1度も呼ばれていなかった＝lib緑でも紙に出ない） */
        + (idx === 0 ? leadBlock() + grandBlock() + carryBlock() : '')
        /* ★2カラム★ 左＝①ご請求の内訳（明細）／右＝②差し引く＋③締め
             ★2段組みは表で作る★（flex だと文が1文字ずつ縦に割れる＝前科あり）。
             ★明細の枠が固定★なので、右の「請求額」は ★中身が1行でも28行でも同じ高さ★に来る。
             紙は A4 の幅で組むので ここは常に2カラム。★狭い画面の1カラムは 入力の画面の側★
             （①→②→③の順番は どちらも同じ＝探す場所が変わらない）。 */
        + (function () {
          /* ★① ご請求の内訳★（左／上）＝明細＋★このブロックの合計★
             見出しの下に線を引く＝★左右の1行目が同じ高さから始まる★（給料明細と同じ） */
          var pageSub = pageLines.reduce(function (a2, x) { return a2 + (Number(x.amount) || 0); }, 0);
          var itemsBlk = '<div class="blk blk-items">'
            + blockHead(caption ? caption : 'ご請求の内訳')
            + '<table class="items"><thead><tr>' + headHtml + '</tr></thead>'
            + '<tbody>' + rowsHtmlOf(pageLines, offset) + '</tbody></table>'
            + (last ? blockSum('明細の合計', yen(tax.subtotal)) : blockSum('このページの小計', yen(pageSub)))
            + '</div>';
          var rightBlk = last ? deductBlock() : '';
          if (layout === 'col1') {
            /* ★1カラム＝上から ①明細 → ②差し引く → ③締め★（順番は2カラムと同じ） */
            return itemsBlk + rightBlk;
          }
          return '<table class="cols2"><tbody><tr>'
            + '<td class="c2-l">' + itemsBlk + '</td>'
            + '<td class="c2-r">' + rightBlk + '</td>'
            + '</tr></tbody></table>';
        })()
        + (last ? totalsBlock() : '')
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
      /* ★角印は薄く重ねる（下の文字を隠し切らない）★
         大きさは会社が決める（10〜40mm・既定21mm）。文字の上に少しかかってよい。 */
      '.seal{display:inline-block;object-fit:contain;margin-top:2mm;opacity:.95;}',
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
      /* ★見出しは切らずに折り返す★
         2カラムにして明細の幅が狭くなり、「数量」が ★「数…」と切れていた★（2026-08-15 実測）。
         ★列の名前は会社が決める＝長い名前も来る★ので、切るのではなく折り返して全部 見せる。 */
      '.items th{background:' + TH.headBg + ';color:' + TH.headInk + ';font-weight:700;font-size:8.5pt;',
      'border:0;padding:1.4mm 1.2mm;line-height:1.35;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',
      '.items td{' + cellBorder + 'padding:1.4mm 1.2mm;vertical-align:top;line-height:1.55;height:' + ROW_H + ';}',
      '.items .c-left{text-align:left;}',
      '.items .c-center{text-align:center;}',
      '.items .c-right{text-align:right;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      '.items .c-wrap{word-break:normal;overflow-wrap:break-word;}',
      '.items .c-nowrap{white-space:nowrap;}',
      '.c-memo{display:block;font-size:8.5pt;color:' + SUB + ';line-height:1.6;margin-top:.5mm;}',
      '.c-empty{text-align:center;color:' + SUB + ';padding:6mm 2mm;}',

      /* ★足元＝左に振込先／右に合計（表の2列＝文が縦に割れない）★ */
      /* ★2カラム（左＝①明細／右＝②控除＋③締め）★
         ★表で組む★（flex だと文が1文字ずつ縦に割れる）。上ぞろえ＝締めの位置が動かない。 */
      '.cols2{width:100%;border-collapse:collapse;table-layout:fixed;margin:0 0 4mm;}',
      '.cols2>tbody>tr>td{vertical-align:top;padding:0;}',
      '.c2-l{width:70%;padding-right:4mm;}',
      '.c2-r{width:30%;}',

      /* ★空の枠★ 罫線は残す（★白黒コピーでも見える濃さ★＝色ではなく濃さで作る） */
      '.r-blank td,.r-blank th{color:transparent;}',
      '.c-blank{background:transparent;}',

      /* ★ブロック（給料明細と同じ作法）★
         見出しの下に線 → 中身（足りない行は高さの決まった空行）→ ★このブロックの合計★（上に線）
         ★左右の行の高さを同じにする★＝同じ番号の行の上端が同じ位置に来る。 */
      '.blk{margin:0 0 4mm;}',
      '.st{font-size:9.5pt;font-weight:700;color:' + ACCENT + ';letter-spacing:.16em;',
      'padding:0 0 1.6mm;border-bottom:0.7pt solid ' + LINE + ';margin:0 0 0;}',
      '.bsum{width:100%;border-collapse:collapse;font-size:9.5pt;margin:0;}',
      '.bsum th{text-align:left;font-weight:700;color:' + INK + ';border:0;border-top:0.9pt solid ' + ACCENT + ';',
      'padding:1.8mm 1.2mm;white-space:nowrap;}',
      '.bsum td{text-align:right;font-weight:700;color:' + INK + ';border:0;border-top:0.9pt solid ' + ACCENT + ';',
      "padding:1.8mm 1.2mm;white-space:nowrap;font-family:'DM Mono',ui-monospace,monospace;}",

      /* ★② 差し引く（控除）★ ★行の高さは明細と同じ★（左右の罫線をそろえる） */
      '.ded{width:100%;border-collapse:collapse;font-size:9.5pt;table-layout:fixed;}',
      '.ded th{text-align:left;font-weight:400;color:' + INK + ';' + cellBorder + 'padding:1.4mm 1.2mm;',
      'height:' + ROW_H + ';line-height:1.55;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',
      '.ded td{text-align:right;white-space:nowrap;' + cellBorder + 'padding:1.4mm 1.2mm;',
      'height:' + ROW_H + ';line-height:1.55;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      /* 右の見出しの行＝左の明細の見出しと同じ高さ・同じ見た目 */
      '.ded-hd th,.ded-hd td{background:' + TH.headBg + ';color:' + TH.headInk + ';font-weight:700;',
      'font-size:8.5pt;border:0;padding:1.4mm 1.2mm;line-height:1.35;height:auto;}',
      '.ded-hd td{text-align:right;font-family:inherit;}',
      '.ded .r-blank th,.ded .r-blank td{color:transparent;}',

      '.foot{width:100%;border-collapse:collapse;table-layout:fixed;margin:0;}',
      '.foot td{vertical-align:top;padding:0;}',
      '.foot-l{width:52%;min-width:70mm;padding-right:6mm;}',
      '.foot-r{width:48%;min-width:70mm;}',

      /* ★小計/消費税/合計＝右下・枠なし・合計の上に線★ */
      '.sums{border-collapse:collapse;font-size:9.5pt;width:100%;margin:0 0 4mm;}',
      '.sums th{text-align:left;color:' + SUB + ';font-weight:400;border:0;',
      'padding:1.4mm 3mm;white-space:nowrap;}',
      '.sums td{text-align:right;border:0;padding:1.4mm 3mm;white-space:nowrap;',
      "font-family:'DM Mono',ui-monospace,monospace;}",
      /* ★紙の中で一番 大きい金額は 頭の「ご請求金額」1つだけ★（給料明細の差引支給額と同じ）。
         締めの中は ★全部 小さく★＝どれを振り込むのか迷わせない。 */
      '.sums-g th{border-top:0.9pt solid ' + ACCENT + ';font-size:12pt;font-weight:700;color:' + INK + ';}',
      '.sums-g td{border-top:0.9pt solid ' + ACCENT + ';font-size:14pt;font-weight:700;color:' + TH.grandInk + ';}',
      '.sums-mid th,.sums-mid td{border-top:0.7pt solid ' + LINE + ';font-weight:700;color:' + INK + ';}',

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

      /* ── 領収書だけの物 ──────────────────────────────────
         ★文が入る所に flex を使わない★（1文字ずつ縦に割れる事故を作らない）＝
         但し書きは「ラベル＋本文」を ★inline で並べ、本文は折り返せる★ 形にする。 */
      '.rc-but{margin:0 0 4mm;font-size:11pt;line-height:1.9;',
      'white-space:normal;word-break:normal;overflow-wrap:break-word;}',
      '.rc-but-l{color:' + SUB + ';white-space:nowrap;}',
      '.rc-but-b{border-bottom:.5pt solid ' + LINE + ';padding:0 1mm .6mm;}',
      '.rc-ack{margin:0 0 6mm;font-size:11pt;color:' + INK + ';}',
      '.rc-sums{margin:0 0 4mm;}',
      '.rc-way{margin:0 0 3mm;font-size:9.5pt;color:' + SUB + ';}',
      '.rc-stamp{display:block;width:100%;min-width:60mm;margin:6mm 0 0;font-size:8.5pt;color:' + SUB + ';',
      'line-height:1.9;white-space:normal;word-break:normal;overflow-wrap:break-word;}',

      '@page{size:A4;margin:0;}',
      '@media print{.sheet{width:210mm;min-width:210mm;padding:14mm 12mm;}}',
    ].join('');
  }

  return {
    build: build, css: css, esc: esc, yen: yen, comma: comma,
    dateStr: dateStr, jpDate: jpDate, honorOf: honorOf, taxLabel: taxLabel,
    paginate: paginate, sealMm: sealMm, TEMPLATE_ID: TEMPLATE_ID,
    ROWS_FIRST: ROWS_FIRST, ROWS_REST: ROWS_REST,
  };
});
