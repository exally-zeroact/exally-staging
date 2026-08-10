/* app.js — 請求書の画面の配線
 * ==============================================================================
 * ★決まりはここに書かない★
 *   消費税 = seikyu/lib/seikyu-tax.js ／ 番号・凍結・入金の状態 = seikyu/lib/seikyu-doc.js
 *   紙      = seikyu/lib/seikyu-paper.js ／ Excelの中身 = seikyu/lib/seikyu-aoa.js
 *   ファイル名 = seikyu/lib/seikyu-name.js ／ 倉庫 = seikyu/js/seikyu-store.js
 *   ここは「押した時に、どれを呼ぶか」だけ。
 *
 * ★税率の数字を1つも書かない★ … 選択肢は SeikyuTax.rates()（唯一の正から）で作る。
 * ★取れなかったを 0 や空にしない★ … 入金が読めない時は「未確認」と出す（0件と作り分ける）。
 * ★押せない時は、押せない理由を出す★ … 黙って無反応にしない。
 */
(function (global) {
  'use strict';

  var DOC = global.SeikyuDoc, TAX = global.SeikyuTax, PAPER = global.SeikyuPaper;
  var NAME = global.SeikyuName, AOA = global.SeikyuAoa, OUT = global.SeikyuOut;

  var TEMPLATE_ID = 'std1';

  var S = {
    sb: null, store: null, suite: null,
    org: null, partners: [], invoices: [], receipts: null,
    cur: null,            // 今 開いている1通（画面の下書き）
    fil: 'all',
    dirty: false,
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) { return PAPER.esc(s); }
  function show(el, on) { if (el) el.style.display = on ? '' : 'none'; }
  function setText(id, t) { var e = $(id); if (e) e.textContent = t || ''; }
  function box(id, text) { var e = $(id); if (!e) return; e.textContent = text || ''; show(e, !!text); }

  function yen(v) { return PAPER.yen(v); }
  function todayYmd() {
    var d = new Date(), z = function (n) { return ('0' + n).slice(-2); };
    return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
  }

  /* ═══ 画面の切り替え ═══ */
  function goScreen(id) {
    ['scr-list', 'scr-edit', 'scr-set'].forEach(function (s) {
      var el = $(s); if (el) el.classList.toggle('active', s === id);
    });
    Array.prototype.forEach.call(document.querySelectorAll('.bn'), function (b) {
      b.classList.toggle('on', b.getAttribute('data-scr') === id);
    });
    try { global.scrollTo(0, 0); } catch (e) { /* 端末によっては動かないが害はない */ }
  }

  /* ═══ 設定（自社の data に同居させる） ═══ */
  function settings() {
    var d = S.org || {};
    var nb = (d.numbering && d.numbering.invoice) || {};
    return {
      format: nb.format || 'ym-seq',
      resetYearly: nb.resetYearly === undefined ? true : !!nb.resetYearly,
      taxMode: d.taxMode === 'inclusive' ? 'inclusive' : 'exclusive',
      rounding: (['floor', 'ceil', 'round'].indexOf(d.taxRounding) >= 0) ? d.taxRounding : 'floor',
      bank: d.bank || '',
    };
  }

  var ROUND_LABEL = { floor: '切り捨て', ceil: '切り上げ', round: '四捨五入' };

  function fillSelect(el, items, value) {
    if (!el) return;
    el.innerHTML = items.map(function (it) {
      return '<option value="' + esc(it.v) + '">' + esc(it.t) + '</option>';
    }).join('');
    if (value !== undefined && value !== null) el.value = value;
  }

  /* 税率の選択肢。★数字は書かない＝唯一の正から作る★ */
  function rateOptions() {
    var rs = TAX.rates() || [];
    return rs.map(function (p) {
      return { v: String(p), t: (Number(p) === 0 ? '対象外' : p + '%') };
    });
  }

  /* ═══ 取引先 ═══ */
  function partnerById(id) {
    for (var i = 0; i < S.partners.length; i++) if (S.partners[i].id === id) return S.partners[i];
    return null;
  }
  function partnerName(inv) {
    var byId = {};
    S.partners.forEach(function (p) { byId[p.id] = p; });
    return DOC.partnerNameOf(inv, byId);
  }

  /* ═══ 一覧 ═══ */
  function payLabel(inv) {
    var st = DOC.paymentStateOf({ id: inv.id, grand_total: (inv.totals && inv.totals.grandTotal) || 0 },
      S.receipts === null ? null : S.receipts);
    return DOC.PAY_STATE_LABEL[st.state] || '';
  }

  function renderList() {
    var host = $('list-body'); if (!host) return;
    var rows = S.invoices.filter(function (v) { return S.fil === 'all' || v.status === S.fil; });
    if (!rows.length) {
      host.innerHTML = '<div class="card"><div class="empty">'
        + (S.invoices.length ? 'この絞り込みに当てはまる請求書はありません。' : 'まだ請求書がありません。「＋ 新しく作る」から作れます。')
        + '</div></div>';
      return;
    }
    host.innerHTML = rows.map(function (v) {
      var tag = v.status === 'issued' ? '<span class="tag tag-issued">発行済</span>'
        : v.status === 'void' ? '<span class="tag tag-void">取り消し</span>'
          : '<span class="tag tag-draft">下書き</span>';
      var g = (v.totals && v.totals.grandTotal);
      return '<button class="iv-row" type="button" data-open="' + esc(v.id) + '">'
        + '<span class="iv-top">' + tag
        + '<span class="iv-no">' + (esc(v.no) || '（未採番）') + '</span>'
        + '<span class="iv-name">' + esc(partnerName(v)) + '</span></span>'
        + '<span class="iv-sub">' + esc(v.issue_ymd || '請求日なし')
        + (v.due_ymd ? '　期限 ' + esc(v.due_ymd) : '')
        + '　' + esc(payLabel(v)) + '</span>'
        + '<span class="iv-sub"><span class="iv-amt">' + (g === undefined || g === null ? '—' : yen(g) + ' 円') + '</span></span>'
        + '</button>';
    }).join('');
    Array.prototype.forEach.call(host.querySelectorAll('[data-open]'), function (b) {
      b.onclick = function () { openInvoice(b.getAttribute('data-open')); };
    });
  }

  function loadList() {
    box('list-err', '');
    return Promise.all([
      S.store.invoices.list('invoice'),
      S.store.receipts.list(),
    ]).then(function (r) {
      S.invoices = r[0] || [];
      S.receipts = r[1];                     // ★null のまま持つ＝「未確認」と「0件」を作り分ける
      renderList();
    }).catch(function (e) {
      box('list-err', '請求書の一覧が読めませんでした（' + ((e && e.message) || 'error') + '）。ネットの状態を確かめて「読み直す」を押してください。');
    });
  }

  /* ═══ 1通の編集 ═══ */
  function blankInvoice() {
    var s = settings();
    return {
      id: '', doc_type: 'invoice', no: '', partner_id: '',
      issue_ymd: todayYmd(), due_ymd: '',
      status: 'draft', tax_mode: s.taxMode, rounding: s.rounding,
      lines: [{ name: '', qty: '', unit: '', price: '', amount: '', rate: firstRate(), memo: '' }],
      totals: {}, snapshot: {}, data: { subject: '', memo: '', noMode: 'auto', term: { kind: 'none', n: 0 } },
      template_id: TEMPLATE_ID, quote_from: '',
    };
  }
  function firstRate() {
    var rs = TAX.rates() || [];
    return rs.length ? String(rs[0]) : '';
  }

  function newInvoice() {
    S.cur = blankInvoice();
    // 取引先が1社だけなら選んでおく（白紙を埋めさせない）
    if (S.partners.length === 1) S.cur.partner_id = S.partners[0].id;
    applyPartnerDefaults();
    fillEdit();
    autoNumber();
    goScreen('scr-edit');
  }

  function openInvoice(id) {
    var v = null;
    for (var i = 0; i < S.invoices.length; i++) if (S.invoices[i].id === id) v = S.invoices[i];
    if (!v) { box('list-err', 'この請求書が見つかりませんでした。「読み直す」を押してください。'); return; }
    S.cur = JSON.parse(JSON.stringify(v));
    if (!S.cur.data) S.cur.data = {};
    if (!Array.isArray(S.cur.lines) || !S.cur.lines.length) S.cur.lines = [{ name: '', qty: '', unit: '', price: '', amount: '', rate: firstRate(), memo: '' }];
    fillEdit();
    goScreen('scr-edit');
  }

  function locked() { return S.cur && !DOC.canEdit(S.cur); }

  function fillEdit() {
    var v = S.cur; if (!v) return;
    setText('edit-h', v.id ? ((v.no || '（未採番）') + '　' + (v.status === 'issued' ? '発行済' : v.status === 'void' ? '取り消し済' : '下書き')) : '新しい請求書');
    show($('edit-locked'), locked());
    // ★別の1通に切り替えたら、前の紙の下見は消す（違う請求書の紙を出したままにしない）
    show($('pv-wrap'), false);

    fillSelect($('e-partner'), [{ v: '', t: '（選んでください）' }].concat(S.partners.map(function (p) {
      return { v: p.id, t: (p.data && p.data.name) || '(名称未設定)' };
    })), v.partner_id || '');
    show($('e-partner-hint'), !S.partners.length ? true : true);
    if (!S.partners.length) {
      setText('e-partner-hint', '取引先がまだ1社もありません。Exally のハブ（共有データ ▸ 取引先）で追加してください。');
    }

    $('e-issue').value = v.issue_ymd || '';
    var term = (v.data && v.data.term) || { kind: 'none', n: 0 };
    fillSelect($('e-term'), DOC.PAY_TERMS.map(function (t) { return { v: t.key, t: t.label }; }), term.kind);
    $('e-termn').value = term.n || '';
    show($('e-termn'), term.kind === 'days' || term.kind === 'nextDay');
    $('e-due').value = v.due_ymd || '';

    Array.prototype.forEach.call($('e-taxmode').querySelectorAll('[data-tm]'), function (b) {
      b.classList.toggle('on', b.getAttribute('data-tm') === v.tax_mode);
    });
    fillSelect($('e-round'), TAX.ROUNDINGS.map(function (k) { return { v: k, t: ROUND_LABEL[k] }; }), v.rounding);

    var nm = (v.data && v.data.noMode) || 'auto';
    Array.prototype.forEach.call($('e-nomode').querySelectorAll('[data-nm]'), function (b) {
      b.classList.toggle('on', b.getAttribute('data-nm') === nm);
    });
    $('e-no').value = v.no || '';
    $('e-subject').value = (v.data && v.data.subject) || '';
    $('e-memo').value = (v.data && v.data.memo) || '';

    renderLines();
    recalc();
    lockInputs();
  }

  /* 発行済み・取り消し済みは触らせない（★押せない理由も出す★） */
  function lockInputs() {
    var ro = locked();
    ['e-partner', 'e-issue', 'e-term', 'e-termn', 'e-due', 'e-round', 'e-no', 'e-subject', 'e-memo'].forEach(function (id) {
      var el = $(id); if (el) el.disabled = ro;
    });
    Array.prototype.forEach.call(document.querySelectorAll('#e-taxmode .seg-b, #e-nomode .seg-b'), function (b) { b.disabled = ro; });
    Array.prototype.forEach.call(document.querySelectorAll('#lines-body input, #lines-body select, #lines-body button'), function (b) { b.disabled = ro; });
    var add = $('b-addline'); if (add) add.disabled = ro;

    var v = S.cur || {};
    $('b-save').disabled = ro;
    $('b-issue').disabled = ro;
    $('b-void').disabled = !DOC.canVoid(v);
    $('b-delete').disabled = !(DOC.canDelete(v) && v.id);

    var why = [];
    if (ro && v.status === 'issued') why.push('発行済みなので、直す・保存する・もう一度発行する はできません（取り消してから作り直します）。');
    if (ro && v.status === 'void') why.push('取り消し済みなので直せません。新しく作り直してください。');
    if (!ro) why.push('取り消しは発行済みのときだけ、削除は下書きのときだけ押せます。');
    if (!v.id && !ro) why.push('まだ保存していないので「削除」は押せません。');
    setText('act-why', why.join(' '));
  }

  /* ── 明細 ── */
  function renderLines() {
    var v = S.cur, host = $('lines-body'); if (!host) return;
    var rates = rateOptions();
    host.innerHTML = v.lines.map(function (ln, i) {
      return '<tr data-i="' + i + '">'
        + '<td class="l-name"><input class="finput" data-f="name" value="' + esc(ln.name || '') + '" placeholder="品名"></td>'
        + '<td class="l-sm"><input class="finput num" data-f="qty" inputmode="decimal" value="' + esc(ln.qty === undefined || ln.qty === null ? '' : ln.qty) + '"></td>'
        + '<td class="l-sm"><input class="finput" data-f="unit" value="' + esc(ln.unit || '') + '" placeholder="式"></td>'
        + '<td class="l-sm"><input class="finput num" data-f="price" inputmode="decimal" value="' + esc(ln.price === undefined || ln.price === null ? '' : ln.price) + '"></td>'
        + '<td class="l-sm"><input class="finput num" data-f="amount" inputmode="numeric" value="' + esc(ln.amount === undefined || ln.amount === null ? '' : ln.amount) + '"></td>'
        + '<td class="l-rate"><select class="finput" data-f="rate">'
        + rates.map(function (r) { return '<option value="' + esc(r.v) + '"' + (String(ln.rate) === r.v ? ' selected' : '') + '>' + esc(r.t) + '</option>'; }).join('')
        + '</select></td>'
        + '<td class="l-x"><button class="l-del" type="button" data-del="' + i + '" aria-label="この行を消す">×</button></td>'
        + '</tr>';
    }).join('');
    Array.prototype.forEach.call(host.querySelectorAll('input,select'), function (el) {
      el.oninput = el.onchange = function () {
        var tr = el.closest('tr'), i = +tr.getAttribute('data-i');
        S.cur.lines[i][el.getAttribute('data-f')] = el.value;
        S.dirty = true;
        recalc();
      };
    });
    Array.prototype.forEach.call(host.querySelectorAll('[data-del]'), function (b) {
      b.onclick = function () {
        var i = +b.getAttribute('data-del');
        S.cur.lines.splice(i, 1);
        if (!S.cur.lines.length) S.cur.lines.push({ name: '', qty: '', unit: '', price: '', amount: '', rate: firstRate(), memo: '' });
        S.dirty = true;
        renderLines(); recalc(); lockInputs();
      };
    });
  }

  /* 画面の文字を、計算に渡せる形に直す（★空は空のまま＝0にしない★） */
  function cleanLines(lines) {
    return (lines || []).map(function (ln) {
      var o = { name: ln.name || '', unit: ln.unit || '', rate: ln.rate, memo: ln.memo || '' };
      if (ln.qty !== '' && ln.qty !== undefined && ln.qty !== null) o.qty = Number(ln.qty);
      if (ln.price !== '' && ln.price !== undefined && ln.price !== null) o.price = Number(ln.price);
      if (ln.amount !== '' && ln.amount !== undefined && ln.amount !== null) o.amount = Number(ln.amount);
      return o;
    }).filter(function (o) {
      // ★何も入っていない行は数えない（空行で「明細0行」の赤を出さない）
      return (o.name || o.qty !== undefined || o.price !== undefined || o.amount !== undefined);
    });
  }

  function currentTax() {
    var v = S.cur;
    return TAX.compute({ lines: cleanLines(v.lines), taxMode: v.tax_mode, rounding: v.rounding });
  }

  function recalc() {
    var v = S.cur; if (!v) return null;
    var t;
    try { t = currentTax(); }
    catch (e) { box('edit-err', (e && e.message) || '計算できませんでした'); return null; }
    var host = $('tot-box');
    if (!t.ok) {
      box('edit-err', t.errors.join('\n'));
      if (host) host.innerHTML = '<div class="hint">合計は、明細が直ったら出ます。</div>';
      return t;
    }
    box('edit-err', '');
    var rows = t.byRate.map(function (b) {
      return '<div class="tot-r"><span class="tot-l">' + esc(b.pct) + '% 対象</span><span class="tot-v">'
        + yen(b.base) + ' 円（消費税 ' + yen(b.tax) + ' 円）</span></div>';
    }).join('');
    if (t.exempt && t.exempt.base) {
      rows += '<div class="tot-r"><span class="tot-l">消費税の対象外</span><span class="tot-v">' + yen(t.exempt.base) + ' 円</span></div>';
    }
    if (host) {
      host.innerHTML = rows
        + '<div class="tot-r"><span class="tot-l">小計</span><span class="tot-v">' + yen(t.subtotal) + ' 円</span></div>'
        + '<div class="tot-r"><span class="tot-l">消費税</span><span class="tot-v">' + yen(t.taxTotal) + ' 円</span></div>'
        + '<div class="tot-r tot-g"><span class="tot-l">合計</span><span class="tot-v">' + yen(t.grandTotal) + ' 円</span></div>';
    }
    return t;
  }

  /* ── 番号 ── */
  function autoNumber() {
    var v = S.cur;
    if (!v || locked()) return Promise.resolve();
    if (((v.data && v.data.noMode) || 'auto') !== 'auto') { setText('e-no-hint', '自分で決めた番号も「使用済み」として数えます。同じ番号は倉庫が受け付けません。'); return Promise.resolve(); }
    var s = settings();
    var p = partnerById(v.partner_id);
    var code = (p && p.data && p.data.code) || '';
    var errs = DOC.validateNumbering({ format: s.format, resetYearly: s.resetYearly, partnerCode: code });
    if (errs.length) {
      $('e-no').value = '';
      v.no = '';
      setText('e-no-hint', errs.join(' '));
      return Promise.resolve();
    }
    return S.store.invoices.usedNos('invoice').then(function (used) {
      var no = DOC.nextNo({ format: s.format, resetYearly: s.resetYearly, ymd: v.issue_ymd, partnerCode: code, existing: used });
      v.no = no;
      $('e-no').value = no;
      setText('e-no-hint', no
        ? '「設定」で決めた形から作りました。同じ番号を二度使わないことは倉庫が守ります。'
        : '番号が作れませんでした（請求日か取引先コードを確かめてください）。');
    }).catch(function () {
      setText('e-no-hint', '使用済みの番号が読めませんでした。番号は手で入れてください。');
    });
  }

  /* ── 支払期限 ── */
  function recalcDue() {
    var v = S.cur; if (!v) return;
    var term = (v.data && v.data.term) || { kind: 'none', n: 0 };
    var d = DOC.dueDateFrom(v.issue_ymd, term);
    if (d) { v.due_ymd = d; $('e-due').value = d; }
  }

  /* 取引先を選んだ時、その相手の「いつまでにもらう約束」を既定にする */
  function applyPartnerDefaults() {
    var v = S.cur; if (!v) return;
    var p = partnerById(v.partner_id);
    var t = p && p.data && p.data.payTerm;
    if (t && t.kind) {
      v.data.term = { kind: t.kind, n: t.n || 0 };
      recalcDue();
    }
  }

  /* ── 紙・Excel ── */
  function paperInput() {
    var v = S.cur;
    var t = recalc();
    if (!t || !t.ok) return null;
    var snap = v.snapshot && v.snapshot.partner ? v.snapshot : null;
    var p = partnerById(v.partner_id);
    var partner = snap ? snap.partner : ((p && p.data) || {});
    var org = snap ? snap.org : (S.org || {});
    if (!snap && S.org) org = Object.assign({}, S.org, { bank: settings().bank });
    var inv = Object.assign({}, v, { lines: cleanLines(v.lines) });
    return { inv: inv, tax: t, partner: partner, org: org };
  }

  function suggestName(ext) {
    var pi = paperInput();
    if (!pi) return null;
    return NAME.suggest({
      docType: pi.inv.doc_type, issueYmd: pi.inv.issue_ymd,
      partnerName: pi.partner.name, grandTotal: pi.tax.grandTotal, ext: ext,
    });
  }

  /* ★落とす前に名前を出して直させる★ */
  var fnPending = null;
  function askName(ext, run) {
    var n = suggestName(ext);
    if (!n) { box('edit-err', '中身がまだ整っていないので、この形では出せません。上の赤い印を直してください。'); return; }
    var base = n.replace(new RegExp('\\.' + ext + '$'), '');
    $('fn-input').value = base;
    setText('fn-ext', '拡張子は . ' + ext + ' が付きます');
    fnPending = { ext: ext, run: run };
    $('fn-ov').classList.add('open');
  }
  function fnClose() { $('fn-ov').classList.remove('open'); fnPending = null; }

  /* ★紙はA4の幅で組んである。スマホの幅にそのまま入れると右が切れて金額が見えない。
       だから紙は縮めず、まるごと縮小して全体を見せる（実物と同じ形のまま小さくする）。 */
  function fitPreview() {
    var f = $('pv'), wrap = $('pv-wrap');
    if (!f || !wrap || wrap.style.display === 'none') return;
    var d = f.contentDocument;
    if (!d || !d.documentElement) return;
    var pw = Math.max(320, d.documentElement.scrollWidth);
    var ph = Math.max(320, d.documentElement.scrollHeight);
    f.style.width = pw + 'px';
    f.style.height = ph + 'px';
    var avail = wrap.clientWidth || pw;
    var k = Math.min(1, avail / pw);
    f.style.transform = 'scale(' + k + ')';
    wrap.style.height = Math.ceil(ph * k) + 'px';
  }

  function doPreview() {
    var pi = paperInput();
    if (!pi) { box('edit-err', '中身がまだ整っていないので、下見が出せません。上の赤い印を直してください。'); return; }
    var built = PAPER.build(pi);
    var f = $('pv');
    f.onload = fitPreview;
    f.srcdoc = built.html;
    show($('pv-wrap'), true);
    // srcdoc は端末によって load が来ないことがあるので、時間でも1度合わせる
    global.setTimeout(fitPreview, 260);
    box('edit-ok', '下の枠が、そのまま刷られる紙です（画面に収まるよう小さくして出しています）。');
  }

  function doPrint(name) {
    var pi = paperInput(); if (!pi) return;
    var built = PAPER.build(Object.assign({}, pi, { title: name }));
    var r = OUT.print(built.html, name);
    if (!r.ok) box('edit-err', r.reason);
    else box('edit-ok', '紙だけの新しい窓を開きました。PDFにする時は、送信先を「PDFに保存」にしてください。');
  }

  function doExcel(name) {
    var pi = paperInput(); if (!pi) return;
    var sheet = AOA.build(pi);
    OUT.excel(sheet, name)
      .then(function () { box('edit-ok', name + ' を保存しました。'); })
      .catch(function (e) { box('edit-err', 'Excelが出せませんでした（' + ((e && e.message) || 'error') + '）'); });
  }

  /* ── 保存・発行 ── */
  function collect() {
    var v = S.cur;
    v.data.subject = $('e-subject').value;
    v.data.memo = $('e-memo').value;
    v.no = $('e-no').value.trim();
    v.issue_ymd = $('e-issue').value || '';
    v.due_ymd = $('e-due').value || '';
    return v;
  }

  function saveDraft() {
    var v = collect();
    var t = recalc();
    v.lines = cleanLines(v.lines);
    v.totals = (t && t.ok) ? {
      subtotal: t.subtotal, taxTotal: t.taxTotal, grandTotal: t.grandTotal,
      byRate: t.byRate, exempt: t.exempt, hasReduced: t.hasReduced,
    } : {};
    if (!v.issue_ymd) { box('edit-err', '請求日を入れてください（下書きでも日付は要ります）。'); return Promise.resolve(); }
    if (!v.no) { box('edit-err', '請求番号が空です。「自動」に戻すか、自分で番号を入れてください。'); return Promise.resolve(); }
    box('edit-err', '');
    return S.store.invoices.saveDraft(v).then(function (r) {
      if (!r.ok) {
        box('edit-err', r.dup
          ? 'この番号（' + v.no + '）は既に使われています。別の番号にしてください。'
          : '保存できませんでした（' + r.reason + '）');
        return;
      }
      S.cur.id = r.id;
      S.dirty = false;
      box('edit-ok', '下書きを保存しました。');
      return loadList().then(function () { fillEdit(); });
    });
  }

  function issue() {
    var v = collect();
    var t = recalc();
    if (!t || !t.ok) { box('edit-err', '計算が通らないので発行できません。上の赤い印を直してください。'); return Promise.resolve(); }
    var p = partnerById(v.partner_id);
    var chk = DOC.validateInvoice({ inv: Object.assign({}, v, { lines: cleanLines(v.lines) }), partner: p, org: { data: S.org || {} } });
    if (!chk.ok) { box('edit-err', chk.errors.join('\n')); box('edit-warn', ''); return Promise.resolve(); }
    box('edit-warn', chk.warnings.join('\n'));
    box('edit-err', '');

    var at = new Date().toISOString();
    var orgData = Object.assign({}, S.org || {}, { bank: settings().bank });
    var snap = DOC.snapshotOf({
      at: at, partner: p, org: { data: orgData }, tax: t, templateId: TEMPLATE_ID,
    });
    var row = Object.assign({}, v, {
      lines: cleanLines(v.lines),
      totals: { subtotal: t.subtotal, taxTotal: t.taxTotal, grandTotal: t.grandTotal, byRate: t.byRate, exempt: t.exempt, hasReduced: t.hasReduced },
      snapshot: snap, template_id: TEMPLATE_ID,
    });
    return S.store.invoices.issue(row, at).then(function (r) {
      if (!r.ok) { box('edit-err', '発行できませんでした（' + r.reason + '）'); return; }
      S.cur.id = r.id; S.cur.no = r.no; S.cur.status = 'issued'; S.cur.issued_at = at;
      S.cur.snapshot = snap; S.cur.totals = row.totals;
      box('edit-ok', '請求書 ' + r.no + ' を発行しました。'
        + (r.bumped ? '（同じ番号が先に使われていたので ' + r.bumped + ' つ進めました）' : '')
        + ' これで中身は固まります。');
      return loadList().then(function () { fillEdit(); });
    });
  }

  function voidIt() {
    var v = S.cur;
    if (!DOC.canVoid(v)) return Promise.resolve();
    if (!global.confirm('請求書 ' + (v.no || '') + ' を取り消しますか？\n（中身は残ります。番号は他で使えません）')) return Promise.resolve();
    return S.store.invoices.voidIt(v.id, new Date().toISOString()).then(function (r) {
      if (!r.ok) { box('edit-err', '取り消せませんでした（' + r.reason + '）'); return; }
      S.cur.status = 'void';
      box('edit-ok', '取り消しました。');
      return loadList().then(function () { fillEdit(); });
    });
  }

  function removeDraft() {
    var v = S.cur;
    if (!(DOC.canDelete(v) && v.id)) return Promise.resolve();
    if (!global.confirm('この下書きを削除しますか？')) return Promise.resolve();
    return S.store.invoices.removeDraft(v.id).then(function (r) {
      if (!r.ok) { box('edit-err', '削除できませんでした（' + r.reason + '）'); return; }
      return loadList().then(function () { goScreen('scr-list'); });
    });
  }

  /* ═══ 設定の画面 ═══ */
  function fillSettings() {
    var s = settings();
    fillSelect($('s-format'), DOC.NUMBER_FORMATS.map(function (f) {
      return { v: f.key, t: f.label + '（' + f.sample + '）' };
    }), s.format);
    $('s-reset').checked = s.resetYearly;
    $('s-taxmode').value = s.taxMode;
    fillSelect($('s-round'), TAX.ROUNDINGS.map(function (k) { return { v: k, t: ROUND_LABEL[k] }; }), s.rounding);
    $('s-bank').value = s.bank;
    settingsHint();

    fillSelect($('s-partner'), [{ v: '', t: '（選んでください）' }].concat(S.partners.map(function (p) {
      return { v: p.id, t: (p.data && p.data.name) || '(名称未設定)' };
    })), '');
    fillSelect($('s-pterm'), DOC.PAY_TERMS.map(function (t) { return { v: t.key, t: t.label }; }), 'none');
    fillPartnerForm('');
  }

  function settingsHint() {
    var f = $('s-format').value, reset = $('s-reset').checked;
    var errs = DOC.validateNumbering({ format: f, resetYearly: reset, partnerCode: 'A001' });
    setText('s-format-hint', errs.length ? errs.join(' ')
      : '「取引先＋年月＋連番」を選ぶ時は、下の欄で取引先コードを入れてください（空のままだと番号を作りません）。');
  }

  function fillPartnerForm(id) {
    var p = partnerById(id);
    var d = (p && p.data) || {};
    $('s-pcode').value = d.code || '';
    $('s-phonor').value = d.honor || d.keisho || '御中';
    $('s-pperson').value = d.person || '';
    $('s-pzip').value = d.zip || '';
    $('s-ptel').value = d.tel || '';
    var t = d.payTerm || { kind: 'none', n: 0 };
    $('s-pterm').value = t.kind || 'none';
    $('s-ptermn').value = t.n || '';
    show($('s-ptermn'), t.kind === 'days' || t.kind === 'nextDay');
    var on = !!p;
    ['s-pcode', 's-phonor', 's-pperson', 's-pzip', 's-ptel', 's-pterm', 's-ptermn'].forEach(function (x) { $(x).disabled = !on; });
    $('b-pt-save').disabled = !on;
  }

  function saveSettings() {
    var patch = {
      numbering: Object.assign({}, (S.org && S.org.numbering) || {}, {
        invoice: { format: $('s-format').value, resetYearly: $('s-reset').checked },
      }),
      taxMode: $('s-taxmode').value,
      taxRounding: $('s-round').value,
      bank: $('s-bank').value,
    };
    var errs = DOC.validateNumbering({ format: patch.numbering.invoice.format, resetYearly: patch.numbering.invoice.resetYearly, partnerCode: 'A001' });
    if (errs.length) { box('set-err', errs.join(' ')); box('set-ok', ''); return Promise.resolve(); }
    box('set-err', '');
    return S.store.org.save(patch).then(function (r) {
      if (!r.ok) { box('set-err', '保存できませんでした（' + r.reason + '）'); return; }
      S.org = r.data;
      box('set-ok', '保存しました。');
    });
  }

  function savePartner() {
    var id = $('s-partner').value;
    if (!id) { box('pt-err', '先に取引先を選んでください。'); return Promise.resolve(); }
    var kind = $('s-pterm').value;
    var add = {
      code: $('s-pcode').value.trim(),
      honor: $('s-phonor').value,
      keisho: $('s-phonor').value,   // ★ハブの取引先画面が読むキー。片方だけ直すと画面で食い違う
      person: $('s-pperson').value.trim(),
      zip: $('s-pzip').value.trim(),
      tel: $('s-ptel').value.trim(),
      payTerm: { kind: kind, n: Math.trunc(Number($('s-ptermn').value) || 0) },
    };
    box('pt-err', '');
    return S.store.partners.patch(id, add).then(function (r) {
      if (!r.ok) { box('pt-err', '保存できませんでした（' + r.reason + '）'); return; }
      box('pt-ok', '保存しました。');
      return S.store.partners.list().then(function (list) { S.partners = list; });
    });
  }

  /* ═══ 配線 ═══ */
  function bind() {
    Array.prototype.forEach.call(document.querySelectorAll('.bn'), function (b) {
      b.onclick = function () {
        var t = b.getAttribute('data-scr');
        if (t === 'scr-edit' && !S.cur) { newInvoice(); return; }
        if (t === 'scr-set') fillSettings();
        goScreen(t);
      };
    });
    Array.prototype.forEach.call(document.querySelectorAll('#fil-seg [data-fil]'), function (b) {
      b.onclick = function () {
        S.fil = b.getAttribute('data-fil');
        Array.prototype.forEach.call(document.querySelectorAll('#fil-seg [data-fil]'), function (x) { x.classList.toggle('on', x === b); });
        renderList();
      };
    });
    $('b-new').onclick = function () { newInvoice(); };
    $('b-reload').onclick = function () { return loadList(); };
    $('b-back').onclick = function () { goScreen('scr-list'); };

    $('e-partner').onchange = function () {
      S.cur.partner_id = $('e-partner').value;
      applyPartnerDefaults();
      return autoNumber();
    };
    $('e-issue').onchange = function () {
      S.cur.issue_ymd = $('e-issue').value;
      recalcDue();
      return autoNumber();
    };
    $('e-term').onchange = function () {
      var k = $('e-term').value;
      S.cur.data.term = { kind: k, n: Math.trunc(Number($('e-termn').value) || 0) };
      show($('e-termn'), k === 'days' || k === 'nextDay');
      recalcDue();
    };
    $('e-termn').oninput = function () {
      S.cur.data.term = { kind: $('e-term').value, n: Math.trunc(Number($('e-termn').value) || 0) };
      recalcDue();
    };
    $('e-due').onchange = function () { S.cur.due_ymd = $('e-due').value; };
    Array.prototype.forEach.call(document.querySelectorAll('#e-taxmode [data-tm]'), function (b) {
      b.onclick = function () {
        S.cur.tax_mode = b.getAttribute('data-tm');
        Array.prototype.forEach.call(document.querySelectorAll('#e-taxmode [data-tm]'), function (x) { x.classList.toggle('on', x === b); });
        recalc();
      };
    });
    $('e-round').onchange = function () { S.cur.rounding = $('e-round').value; recalc(); };
    Array.prototype.forEach.call(document.querySelectorAll('#e-nomode [data-nm]'), function (b) {
      b.onclick = function () {
        var m = b.getAttribute('data-nm');
        S.cur.data.noMode = m;
        Array.prototype.forEach.call(document.querySelectorAll('#e-nomode [data-nm]'), function (x) { x.classList.toggle('on', x === b); });
        if (m === 'auto') return autoNumber();
        setText('e-no-hint', '自分で決めた番号も「使用済み」として数えます。同じ番号は倉庫が受け付けません。');
      };
    });
    $('e-no').oninput = function () { S.cur.no = $('e-no').value; };
    $('e-subject').oninput = function () { S.cur.data.subject = $('e-subject').value; };
    $('e-memo').oninput = function () { S.cur.data.memo = $('e-memo').value; };
    $('b-addline').onclick = function () {
      S.cur.lines.push({ name: '', qty: '', unit: '', price: '', amount: '', rate: firstRate(), memo: '' });
      renderLines(); recalc(); lockInputs();
    };

    $('b-preview').onclick = function () { doPreview(); };
    $('b-print').onclick = function () { askName('pdf', doPrint); };
    $('b-xlsx').onclick = function () { askName('xlsx', doExcel); };
    $('fn-ok').onclick = function () {
      if (!fnPending) return;
      var ext = fnPending.ext, run = fnPending.run;
      var base = NAME.sanitize($('fn-input').value) || 'seikyu';
      fnClose();
      run(base + '.' + ext);
    };
    $('fn-cancel').onclick = fnClose;

    $('b-save').onclick = function () { return saveDraft(); };
    $('b-issue').onclick = function () { return issue(); };
    $('b-void').onclick = function () { return voidIt(); };
    $('b-delete').onclick = function () { return removeDraft(); };

    $('s-format').onchange = settingsHint;
    $('s-reset').onchange = settingsHint;
    $('s-partner').onchange = function () { fillPartnerForm($('s-partner').value); };
    $('s-pterm').onchange = function () {
      var k = $('s-pterm').value;
      show($('s-ptermn'), k === 'days' || k === 'nextDay');
    };
    $('b-set-save').onclick = function () { return saveSettings(); };
    $('b-pt-save').onclick = function () { return savePartner(); };

    // 画面を回した・幅が変わった時も、下見が切れないように合わせ直す
    global.addEventListener('resize', fitPreview);
  }

  /* ログインが済んでから呼ばれる（seikyu/js/auth.js） */
  function attach(sb) {
    S.sb = sb;
    S.suite = global.SuiteData.create({ client: sb });
    S.store = global.SeikyuStore.create({ client: sb, suite: S.suite });
    bind();
    return Promise.all([
      S.suite.org.get().catch(function () { return null; }),
      S.suite.partners.list().catch(function () { return []; }),
    ]).then(function (r) {
      S.org = r[0] || {};
      S.partners = r[1] || [];
      return loadList();
    }).then(function () { return S.store; });
  }

  global.SeikyuApp = {
    attach: attach,
    _state: S,          // テストから中を見るため（画面の外からは使わない）
    _go: goScreen,
    _new: newInvoice,
    _fillSettings: fillSettings,
  };
})(window);
