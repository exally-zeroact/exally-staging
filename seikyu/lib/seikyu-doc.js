/* seikyu-doc.js — 請求書という「物」の決まり（唯一の正）
 * ==============================================================================
 * ★なぜこのファイルが要るのか（代行請求の実物を読んで分かった事）★
 *   代行請求アプリには ★請求書という物が1つも保存されていない★。
 *   紙は毎回、明細から作り直している。その結果:
 *     ・後から明細を1行直すと ★去年 出した請求書の金額が変わる★
 *     ・番号が「会社の並び順」から作られていて ★会社を1社足すと過去の番号がズレる★
 *     ・出した控えが残らない（適格請求書の写しの保存に直結）
 *   → ★発行した時に決めて、行に保存して、固める★。それがこのファイルの仕事。
 *
 * ★番号について（法律のルールは無い＝自由）★
 *   大事なのは1つだけ ＝ ★同じ番号を二度使わない★
 *   重複すると入金の消し込みでどの請求か分からなくなる（二重請求・誤督促）。
 *   ★このlibは「次の番号を出す」「弾かれたら1つ進める」だけ。
 *     最後の砦は倉庫の一意制約(uq_pay_invoices_no)★。画面のチェックだけにしない。
 *   ★インボイスの登録番号(T+13桁)とは まったくの別物★。同じ欄に出さない。
 *
 * ★画面に依らない（DOMを1つも触らない）／時計も乱数も使わない★
 *   ＝素のNodeで全パターン回せる。発行時刻は呼ぶ側が渡す。
 *
 * 【利用】ブラウザ window.SeikyuDoc（先に seikyu-tax.js を読む）
 *         Node    require('./seikyu-doc.js')
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./seikyu-tax.js'));
  else root.SeikyuDoc = factory(root.SeikyuTax);
})(typeof self !== 'undefined' ? self : this, function (TAX) {
  'use strict';
  if (!TAX || !TAX.compute) throw new Error('seikyu-tax.js を先に読んでください');

  var DOC_TYPES = ['invoice', 'quote'];
  var STATUSES = ['draft', 'issued', 'void'];

  /* ★発行したら固まる列★
     倉庫のトリガ(kyuyo.pay_invoices_freeze)と同じ並びでなければ赤になる
     （seikyu/tests/schema-contract.test.mjs が突き合わせる）。
     status/sent_at/deleted_at は入っていない＝取り消し・送った記録は後から入る。 */
  var FROZEN_FIELDS = [
    'doc_type', 'no', 'partner_id', 'issue_ymd', 'due_ymd',
    'tax_mode', 'rounding', 'lines', 'totals', 'snapshot', 'template_id', 'issued_at',
  ];

  /* 番号の形。★「自分で決める(manual)」も許す（司さん指示）★ */
  var NUMBER_FORMATS = [
    { key: 'seq', label: '連番', sample: '00001', digits: 5, scope: 'all' },
    { key: 'y-seq', label: '年＋連番', sample: '2026-0001', digits: 4, scope: 'year' },
    { key: 'ym-seq', label: '年月＋連番', sample: '202609-001', digits: 3, scope: 'month' },
    { key: 'p-ym-seq', label: '取引先＋年月＋連番', sample: 'A001-202609-01', digits: 2, scope: 'partner-month' },
    { key: 'manual', label: '自分で決める', sample: '（自由入力）', digits: 0, scope: 'manual' },
  ];
  function fmtDef(key) {
    for (var i = 0; i < NUMBER_FORMATS.length; i++) if (NUMBER_FORMATS[i].key === key) return NUMBER_FORMATS[i];
    return null;
  }

  /* ── 日付（時差でズレないよう、文字列を自分で解いて UTC で数える） ──────── */
  function parseYmd(ymd) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(ymd || ''));
    if (!m) return null;
    var y = +m[1], mo = +m[2], d = +m[3];
    var t = new Date(Date.UTC(y, mo - 1, d));
    // ★存在しない日（2026-02-30 等）を黙って繰り上げない
    if (t.getUTCFullYear() !== y || t.getUTCMonth() !== mo - 1 || t.getUTCDate() !== d) return null;
    return { y: y, m: mo, d: d };
  }
  function ymdOf(y, m, d) { return y + '-' + String(m).padStart(2, '0') + '-' + String(d).padStart(2, '0'); }
  function eomOf(y, m) { return new Date(Date.UTC(y, m, 0)).getUTCDate(); } // m は1始まり
  function addDays(p, n) {
    var t = new Date(Date.UTC(p.y, p.m - 1, p.d + n));
    return ymdOf(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate());
  }
  function nextMonth(p) { return p.m === 12 ? { y: p.y + 1, m: 1 } : { y: p.y, m: p.m + 1 }; }

  var PAY_TERMS = [
    { key: 'none', label: '決めていない' },
    { key: 'eom', label: '今月末' },
    { key: 'nextEom', label: '翌月末' },
    { key: 'days', label: '◯日後' },
    { key: 'nextDay', label: '翌月◯日' },
  ];
  /** その請求日から「いつまでにもらう約束か」。
      ★決めていない・読めない日付なら空（勝手な期限を作らない）★ */
  function dueDateFrom(ymd, term) {
    var t = term || {};
    var kind = String(t.kind || 'none');
    if (kind === 'none') return '';
    var p = parseYmd(ymd);
    if (!p) return '';
    var n = Math.trunc(Number(t.n) || 0);
    if (kind === 'eom') return ymdOf(p.y, p.m, eomOf(p.y, p.m));
    if (kind === 'nextEom') { var q = nextMonth(p); return ymdOf(q.y, q.m, eomOf(q.y, q.m)); }
    if (kind === 'days') return addDays(p, n);
    if (kind === 'nextDay') {
      var r = nextMonth(p), last = eomOf(r.y, r.m);
      return ymdOf(r.y, r.m, Math.min(Math.max(1, n), last)); // ★その月に無い日は末日に寄せる
    }
    return ''; // 知らない決め方は「決めていない」と同じ扱い
  }

  /* ── 番号 ────────────────────────────────────────────────────── */
  function pad(n, d) { return String(n).padStart(d, '0'); } // ★padStartは切らない＝桁があふれても伸びるだけ

  /** 形とその通し番号から、人が見る番号を作る。作れない時は空（★空欄のまま作らない★） */
  function formatNo(o) {
    o = o || {};
    var def = fmtDef(o.format);
    if (!def || def.key === 'manual') return '';
    var seq = Math.trunc(Number(o.seq) || 0);
    if (seq < 1) return '';
    if (def.key === 'seq') return pad(seq, def.digits);
    var p = parseYmd(o.ymd);
    if (!p) return '';
    var yyyy = String(p.y), ym = yyyy + String(p.m).padStart(2, '0');
    if (def.key === 'y-seq') return yyyy + '-' + pad(seq, def.digits);
    if (def.key === 'ym-seq') return ym + '-' + pad(seq, def.digits);
    var code = String(o.partnerCode || '').trim();
    if (!code) return ''; // ★取引先コードが無いのに '-202609-01' のような番号を作らない
    return code + '-' + ym + '-' + pad(seq, def.digits);
  }

  function esc(s) { return String(s).replace(/[.*+?^${}()|[\]\\-]/g, '\\$&'); }

  /** 使用済みの番号一覧から、次の番号を出す。
   *  existing … ★同じ account × 同じ doc_type の番号を全部★（呼ぶ側が絞って渡す）
   *  resetYearly … 既定 true（年/年月が変われば1に戻る）。false なら通し番号を続ける。
   *  ★自分で決めた番号も「使用済み」として数える＝自動採番がぶつからない★ */
  function nextNo(o) {
    o = o || {};
    var def = fmtDef(o.format);
    if (!def || def.key === 'manual') return '';
    var reset = o.resetYearly === undefined ? true : !!o.resetYearly;
    var list = Array.isArray(o.existing) ? o.existing : [];
    var p = def.key === 'seq' ? null : parseYmd(o.ymd);
    if (def.key !== 'seq' && !p) return '';
    var code = String(o.partnerCode || '').trim();
    if (def.key === 'p-ym-seq' && !code) return '';

    var re;
    if (def.key === 'seq') re = /^(\d+)$/;
    else if (reset) {
      var prefix = def.key === 'y-seq' ? String(p.y) + '-'
        : def.key === 'ym-seq' ? String(p.y) + String(p.m).padStart(2, '0') + '-'
          : code + '-' + String(p.y) + String(p.m).padStart(2, '0') + '-';
      re = new RegExp('^' + esc(prefix) + '(\\d+)$');
    } else {
      // ★「続ける」＝年/年月が変わっても、その形で出した全部の通し番号を見る
      re = def.key === 'y-seq' ? /^\d{4}-(\d+)$/
        : def.key === 'ym-seq' ? /^\d{6}-(\d+)$/
          : new RegExp('^' + esc(code) + '-\\d{6}-(\\d+)$');
    }

    var max = 0;
    for (var i = 0; i < list.length; i++) {
      var m = re.exec(String(list[i] == null ? '' : list[i]));
      if (!m) continue;                       // ★読めない番号は無視するが落ちない
      var v = parseInt(m[1], 10);
      if (Number.isFinite(v) && v > max) max = v;
    }
    return formatNo({ format: def.key, seq: max + 1, ymd: o.ymd, partnerCode: code });
  }

  /** 倉庫の一意制約に弾かれた時、番号を1つ進めて出し直す（同時発行の再試行）。
   *  ★数で終わらない番号は自動で進めない（人が決めた番号を勝手にいじらない）★ */
  function bumpNo(no) {
    var m = /^(.*?)(\d+)$/.exec(String(no || ''));
    if (!m) return '';
    return m[1] + pad(parseInt(m[2], 10) + 1, m[2].length);
  }

  /** 番号の決め方そのものの検査（設定画面で止める） */
  function validateNumbering(o) {
    o = o || {};
    var errs = [];
    var def = fmtDef(o.format);
    if (!def) { errs.push('番号の形が選ばれていません'); return errs; }
    if (def.key === 'seq' && o.resetYearly === true) {
      // ★番号に年が入っていないのに毎年1に戻すと、去年の番号と必ずぶつかる
      errs.push('「連番」だけの形で毎年1に戻すと、去年の番号とぶつかります（年を入れる形を選ぶか、続ける設定にしてください）');
    }
    if (def.key === 'p-ym-seq' && !String(o.partnerCode || '').trim()) {
      errs.push('取引先コードが空です（この形は取引先ごとのコードが要ります）');
    }
    return errs;
  }

  /* ── 発行したら固まる ────────────────────────────────────────── */
  function statusOf(inv) { return (inv && inv.status) || 'draft'; }
  function canEdit(inv) { return statusOf(inv) === 'draft'; }
  function canDelete(inv) { return statusOf(inv) === 'draft'; } // ★発行済みは消せない＝番号を欠番にしない
  function canVoid(inv) { return statusOf(inv) === 'issued'; }

  /** 発行の写し。★紙に出た物をそのまま残す＝あとでマスタを直しても紙は変わらない★
   *  at（発行時刻）は呼ぶ側が渡す＝このlibは時計を持たない。 */
  /* ═══ ★お金の順番は1か所で決める★ ═══
     今回請求額 →（＋繰越）→ 合計請求額 →（−源泉）→ 差引お支払額

     ここを紙・Excel・画面の3か所で別々に書くと、必ずどこかで食い違う。
     2026-08-11 実機で発生：繰越を出しているのに 差引お支払額 が繰越を無視し、
     ★1,111,000 を請求しながら 997,900 と書いた★（11,000 少なく振り込まれる）。

     ・繰越の入金が読めていない時は合計が出せない → ★null を返す（0にしない）★
     ・源泉は「この1通の報酬」に掛ける（繰越は前回すでに掛けてある）ので、
       引く相手は 合計請求額。掛ける相手は 今回の報酬。ここを混ぜない。 */

  /** 実際に請求している額（繰越があれば足したあと）。読めない時は null */
  function billedOf(tax, carry) {
    var g = tax && Number(tax.grandTotal);
    if (!Number.isFinite(g)) return null;
    if (!carry) return g;
    if (carry.grandTotal === null || carry.grandTotal === undefined) return null;  // 入金が未確認
    return Number(carry.grandTotal);
  }

  /** 実際に振り込んでもらう額（源泉があれば引いたあと）。読めない時は null */
  function payableOf(tax, carry, gensen) {
    var billed = billedOf(tax, carry);
    if (billed === null) return null;
    if (!gensen || !gensen.on) return billed;
    var a = Number(gensen.amount);
    if (!Number.isFinite(a)) return null;
    return billed - a;
  }

  function snapshotOf(o) {
    o = o || {};
    var pd = (o.partner && o.partner.data) || {};
    var od = (o.org && o.org.data) || {};
    var t = o.tax || {};
    return {
      at: o.at || '',
      partner: {
        id: (o.partner && o.partner.id) || '',
        // ★敬称は hub の取引先画面が既に `keisho` で保存している（js/hub.js savePt）。
        //   `honor` しか見ないと、人が選んだ「様」が紙で「御中」に化ける（②で実物と突き合わせて判明）。
        //   契約の名前は `honor` のまま。読む時だけ既存の `keisho` も受ける。
        name: pd.name || '', honor: pd.honor || pd.keisho || '御中', person: pd.person || '',
        zip: pd.zip || '', addr: pd.addr || '', tel: pd.tel || '',
        invoiceNo: pd.invoiceNo || '', code: pd.code || '',
      },
      org: {
        yago: od.yago || '', addr: od.addr || '', tel: od.tel || '',
        invoiceNo: od.invoiceNo || '', bank: od.bank || '',
        sealDataUrl: od.sealDataUrl || '', sealSizeMm: sealSizeMm(od.sealSizeMm), logoDataUrl: od.logoDataUrl || '',
      },
      totals: { subtotal: t.subtotal || 0, taxTotal: t.taxTotal || 0, grandTotal: t.grandTotal || 0 },
      byRate: t.byRate || [],
      exempt: t.exempt || { base: 0 },
      hasReduced: !!t.hasReduced,
      templateId: o.templateId || '',
      // ★どの列で刷ったかも写しに残す★
      //   snapshot は発行後に固まる列なので、あとで会社が列を足しても
      //   出した紙は そのままの並びで刷り直せる（紙と控えが食い違わない）。
      cols: o.cols ? {
        items: (o.cols.items || []).slice(),
        widths: Object.assign({}, o.cols.widths || {}),
        aligns: Object.assign({}, o.cols.aligns || {}),
      } : null,
    };
  }

  /** 一覧に出す取引先名。★取れなかったを空欄にしない★
   *  発行済み＝写しの名前（紙と同じ物）／下書き＝マスタの名前。 */
  function partnerNameOf(inv, partnersById) {
    var snapName = inv && inv.snapshot && inv.snapshot.partner && inv.snapshot.partner.name;
    if (snapName) return snapName;
    var pid = (inv && inv.partner_id) || '';
    if (!pid) return '（取引先が未選択）';
    var p = (partnersById || {})[pid];
    var nm = p && p.data && p.data.name;
    if (nm) return nm;
    return '（消えた取引先）';
  }

  /* ── 角印（会社の印） ──────────────────────────────────────────
     ★紙に重ねるだけの飾りではない★。押してある/無いで相手の受け取り方が変わるので、
     「入れられる」「大きさを変えられる」「消せる」を揃える。
     代行請求の実物と同じ扱い（invoice-pdf.js:927 hankoSizeMm 既定21mm・薄く重ねる）。 */
  var SEAL_DEFAULT_MM = 21;
  var SEAL_MIN_MM = 10;
  var SEAL_MAX_MM = 40;
  /* 倉庫の1行に画像を入れるので上限を決める。★超えたら黙って縮めずに赤で返す★
     （黙って縮めると「押したはずの印が薄い/欠ける」になり、押した本人が気づけない） */
  var SEAL_MAX_BYTES = 300 * 1024;

  function sealSizeMm(v) {
    var n = Number(v);
    if (!Number.isFinite(n)) return SEAL_DEFAULT_MM;
    return Math.max(SEAL_MIN_MM, Math.min(SEAL_MAX_MM, Math.round(n)));
  }

  /** 角印に使える画像か。返り = { ok, reason }
   *  ・png / jpeg の data URL だけ（外のURLは受けない＝紙を刷る時に読めないと空白になる）
   *  ・上限を超えたら赤（何KBかを言う） */
  function validateSeal(dataUrl, opts) {
    var max = (opts && opts.maxBytes) || SEAL_MAX_BYTES;
    var s = String(dataUrl || '');
    if (!s) return { ok: false, reason: '画像が選ばれていません' };
    var m = /^data:image\/(png|jpe?g);base64,([A-Za-z0-9+/=]+)$/.exec(s);
    if (!m) return { ok: false, reason: '角印に使えるのは PNG か JPEG の画像だけです' };
    // base64 の長さから元の大きさを出す（末尾の = は数えない）
    var b64 = m[2];
    var pad = (b64.slice(-2) === '==') ? 2 : (b64.slice(-1) === '=' ? 1 : 0);
    var bytes = Math.floor(b64.length * 3 / 4) - pad;
    if (bytes > max) {
      return { ok: false, bytes: bytes, reason: '画像が大きすぎます（' + Math.round(bytes / 1024) + 'KB）。'
        + Math.round(max / 1024) + 'KB までにしてください（角印は小さく写れば十分です）' };
    }
    return { ok: true, bytes: bytes };
  }

  /* ── 入金 ────────────────────────────────────────────────────── */
  /** 請求1通の入金の状態。
   *  ★receipts が null ＝「取れなかった(未確認)」。0件と必ず違う答えを返す★
   *  ★過入金を0でクランプしない（多く入った事実を残す）★ */
  function paymentStateOf(inv, receipts) {
    var total = Number((inv && inv.grand_total) || 0);
    if (receipts == null) {
      return { state: 'unknown', paid: null, remain: null, count: null, lastYmd: '', total: total };
    }
    var id = (inv && inv.id) || null;
    var paid = 0, count = 0, lastYmd = '';
    for (var i = 0; i < receipts.length; i++) {
      var r = receipts[i] || {};
      if (r.deleted_at) continue;                                   // 消した入金は数えない
      if (id && r.invoice_id !== undefined && r.invoice_id !== id) continue; // 別の請求の入金
      paid += Number(r.amount) || 0;
      count++;
      if (r.ymd && r.ymd > lastYmd) lastYmd = r.ymd;
    }
    var remain = total - paid;
    var state = paid === 0 ? 'unpaid' : remain > 0 ? 'partial' : remain === 0 ? 'paid' : 'over';
    return { state: state, paid: paid, remain: remain, count: count, lastYmd: lastYmd, total: total };
  }
  var PAY_STATE_LABEL = {
    unknown: '入金は未確認', unpaid: '未入金', partial: '一部入金', paid: '入金済', over: '過入金',
  };

  /* ── 入金を「記録する」側の決まり ────────────────────────────
     ★1回＝1行で足す。上書きしない★
       代行請求は `PAYMENTS["会社::月"] = {paid,paidDate,note}` ＝ 1請求に1行しか持てない。
       2回に分けて払われると、2回目で1回目が消える（分割払いの履歴が残らない）。
       うちは入金1回につき1行を足すだけ＝分割払いも過入金も返金も、そのまま履歴に残る。
     ★充てた結果は保存しない★（飲み屋 nomiya-core.js と同じ考え方）
       残り・過入金は毎回その場で数える（paymentStateOf）。正が1つで済み、
       記録を消せば数字もその場でやり直される。
     ★0円は記録させない★＝倉庫の check (amount <> 0) と同じ言葉でここでも断る。
       「入っていない」と「0円 入った」を作り分けると、督促の判断ができなくなる。 */
  var PAY_METHODS = ['振込', '現金', '相殺', '手形', 'その他'];

  /** 画面に打たれた金額を円の整数に直す。
   *  ★読めなかったら null（0にしない）★ … 0にすると「0円の入金」として記録できてしまう。
   *  返金はマイナスで入れる（−2,000 など）。 */
  function receiptAmountOf(v) {
    var raw = String(v == null ? '' : v).trim().replace(/[,\s]/g, '');
    if (!raw) return null;
    if (!/^-?\d+$/.test(raw)) return null;      // 小数・文字・全角は受けない（1円単位）
    var n = Number(raw);
    return Number.isFinite(n) ? n : null;
  }

  /** 入金1件が記録できる形か。返り = { ok, errors:[] }
   *  ★押す前（ボタンの中の理由）と押した後（赤い印）で同じ判定を使う★
   *  ＝同じ状態を2か所で別々に書かない。 */
  function validateReceipt(o) {
    o = o || {};
    var errors = [];
    if (!parseYmd(o.ymd)) errors.push('入金日を入れてください');
    var raw = String(o.amount == null ? '' : o.amount).trim();
    if (!raw) errors.push('金額を入れてください');
    else {
      var n = receiptAmountOf(raw);
      if (n === null) errors.push('金額は1円単位の数字で入れてください');
      else if (n === 0) errors.push('0円は記録できません（返金はマイナスで入れてください）');
    }
    if (String(o.memo || '').length > 200) errors.push('備考が長すぎます（200文字まで）');
    if (String(o.method || '').length > 20) errors.push('方法が長すぎます（20文字まで）');
    return { ok: errors.length === 0, errors: errors };
  }

  /* ── 発行前の検査（空欄のまま出させない） ──────────────────────── */
  function validateInvoice(o) {
    o = o || {};
    var inv = o.inv || {};
    var errors = [], warnings = [];

    if (DOC_TYPES.indexOf(inv.doc_type || 'invoice') < 0) errors.push('書類の種類が不明です（' + inv.doc_type + '）');
    if (!String(inv.no || '').trim()) errors.push('請求番号が空です');
    if (!String(inv.partner_id || '').trim()) errors.push('取引先が選ばれていません');
    else if (!o.partner) errors.push('取引先がマスタに見つかりません（消された可能性があります）');
    var ip = parseYmd(inv.issue_ymd);
    if (!ip) errors.push('請求日が入っていません（または読めません）');
    if (inv.due_ymd) {
      var dp = parseYmd(inv.due_ymd);
      if (!dp) errors.push('支払期限が読めません');
      else if (ip && inv.due_ymd < inv.issue_ymd) errors.push('支払期限が請求日より前になっています');
    }

    var lines = Array.isArray(inv.lines) ? inv.lines : [];
    if (!lines.length) errors.push('明細が1行もありません');

    var tax = TAX.compute({ lines: lines, taxMode: inv.tax_mode, rounding: inv.rounding });
    if (!tax.ok) for (var i = 0; i < tax.errors.length; i++) errors.push(tax.errors[i]);

    if (tax.ok && lines.length) {
      if (tax.grandTotal === 0) errors.push('合計が0円です（0円の請求書は出せません）');
      else if (tax.grandTotal < 0) warnings.push('合計がマイナスです（返金の請求で間違いないか確かめてください）');
    }

    var od = (o.org && o.org.data) || {};
    if (!String(od.invoiceNo || '').trim()) {
      warnings.push('自社の登録番号（インボイス）が空です。このままでも請求書は出せますが、適格請求書にはなりません');
    }
    if (tax.ok && tax.hasReduced && !String(od.invoiceNo || '').trim()) {
      warnings.push('軽減税率の行がありますが、登録番号が無いため税率ごとの区分は参考表示になります');
    }

    return { ok: errors.length === 0, errors: errors, warnings: warnings, tax: tax };
  }

  /** 見積 → 請求 に変換する形（★番号は請求の系列で採り直す★） */
  function convertQuoteToInvoice(q) {
    q = q || {};
    return {
      doc_type: 'invoice',
      quote_from: q.id || '',
      status: 'draft',
      no: '',                       // ★見積の番号を持ち込まない
      partner_id: q.partner_id || '',
      issue_ymd: '', due_ymd: '',
      tax_mode: q.tax_mode, rounding: q.rounding,
      lines: (q.lines || []).slice(),
      totals: {}, snapshot: {},
      template_id: q.template_id || '',
    };
  }

  return {
    DOC_TYPES: DOC_TYPES, STATUSES: STATUSES, FROZEN_FIELDS: FROZEN_FIELDS,
    NUMBER_FORMATS: NUMBER_FORMATS, PAY_TERMS: PAY_TERMS, PAY_STATE_LABEL: PAY_STATE_LABEL,
    PAY_METHODS: PAY_METHODS, receiptAmountOf: receiptAmountOf, validateReceipt: validateReceipt,
    formatNo: formatNo, nextNo: nextNo, bumpNo: bumpNo, validateNumbering: validateNumbering,
    dueDateFrom: dueDateFrom, parseYmd: parseYmd,
    canEdit: canEdit, canDelete: canDelete, canVoid: canVoid, statusOf: statusOf,
    billedOf: billedOf, payableOf: payableOf,
    snapshotOf: snapshotOf, partnerNameOf: partnerNameOf,
    validateSeal: validateSeal, sealSizeMm: sealSizeMm,
    SEAL_DEFAULT_MM: SEAL_DEFAULT_MM, SEAL_MIN_MM: SEAL_MIN_MM, SEAL_MAX_MM: SEAL_MAX_MM, SEAL_MAX_BYTES: SEAL_MAX_BYTES,
    paymentStateOf: paymentStateOf,
    validateInvoice: validateInvoice, convertQuoteToInvoice: convertQuoteToInvoice,
  };
});
