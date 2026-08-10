/* seikyu-store.js — 請求書の倉庫（kyuyo.pay_invoices / kyuyo.pay_receipts）
 * ==============================================================================
 * 契約の一次情報 = docs/SPEC_seikyu_data.md ／ 棚 = supabase/schema-seikyu.sql
 *
 * ★取引先と自社は新しく作らない★
 *   既にある pay_partners / pay_org を js/suite-data.js（E0 共有データ層）越しに使う。
 *   ここで別の取引先マスタを持つと二重管理になる。
 *
 * ★最後の砦は倉庫★
 *   番号の重複は uq_pay_invoices_no が止める。画面のチェックだけを信じない。
 *   弾かれたら seikyu-doc.bumpNo() で1つ進めて出し直す（同時発行の再試行）。
 *
 * ★失敗を握り潰さない★
 *   書けなかったら {ok:false, reason} を返す。嘘の成功を返さない。
 *
 * 【利用】window.SeikyuStore.create({ client, suite })
 */
(function (global) {
  'use strict';

  var MAX_BUMP = 20;   // 同時発行のぶつかりで番号を進める上限（これを超えたら人に返す）

  function err(e) { return (e && (e.message || e.reason)) || 'error'; }
  /* 一意制約に弾かれたか（番号の重複）。PostgREST は 23505 を返す */
  function isDup(e) {
    var s = String((e && (e.code || e.message)) || '');
    return /23505/.test(s) || /duplicate key|uq_pay_invoices_no/i.test(String((e && e.message) || ''));
  }

  function create(opts) {
    opts = opts || {};
    var sb = opts.client;
    var suite = opts.suite || null;
    if (!sb) throw new Error('SeikyuStore.create: client(supabaseクライアント) が必要です');
    var DOC = global.SeikyuDoc;
    if (!DOC) throw new Error('seikyu-doc.js を先に読んでください');

    var seq = 0;
    var newId = opts.newId || function (p) {
      seq++;
      return p + '_' + Math.abs(Date.now()).toString(36) + '_' + seq.toString(36) + Math.random().toString(36).slice(2, 6);
    };
    var nowIso = opts.now || function () { return new Date().toISOString(); };

    function uid() {
      return Promise.resolve(sb.auth.getUser())
        .then(function (r) { return (r && r.data && r.data.user && r.data.user.id) || null; })
        .catch(function () { return null; });
    }

    /* 全件ページング（PostgREST 既定 max_rows=1000 で"黙って"切れるのを根治）。
       ★請求の一覧が静かに過少になると、去年の請求が消えたように見える★ */
    function fetchAllQ(build) {
      var out = [], from = 0, size = 1000;
      function step() {
        return Promise.resolve(build(from, from + size - 1)).then(function (r) {
          if (r && r.error) return { data: null, error: r.error };
          var got = (r && r.data) || [];
          out = out.concat(got);
          if (!got.length || r.count == null || out.length >= r.count) {
            return { data: out, error: null, count: (r && r.count != null) ? r.count : out.length };
          }
          from += got.length;
          return step();
        });
      }
      return step();
    }

    var COLS = 'id,doc_type,no,partner_id,issue_ymd,due_ymd,status,tax_mode,rounding,'
      + 'lines,totals,snapshot,data,template_id,quote_from,issued_at,sent_at,voided_at,updated_at';

    var api = {
      /* ═══ 請求書 ═══ */
      invoices: {
        // 一覧（消していない物だけ・新しい順）。docType 省略で請求書
        list: function (docType) {
          var dt = docType || 'invoice';
          return fetchAllQ(function (a, b) {
            return sb.from('pay_invoices').select(COLS, { count: 'exact' })
              .eq('doc_type', dt).is('deleted_at', null)
              /* ★並びをはっきり決める＝新しい順（請求日）→ 同じ日は番号の大きい順 → 最後はidで固定★
                 （最後まで決めないと、同じ日・同じ番号の時に開くたび順が変わる） */
              .order('issue_ymd', { ascending: false })
              .order('no', { ascending: false })
              .order('id', { ascending: false })
              .range(a, b);
          }).then(function (r) {
            if (r && r.error) throw new Error(err(r.error));
            return r.data || [];
          });
        },
        get: function (id) {
          return Promise.resolve(sb.from('pay_invoices').select(COLS).eq('id', id).maybeSingle())
            .then(function (r) {
              if (r && r.error) throw new Error(err(r.error));
              return (r && r.data) || null;
            });
        },
        /* 使用済みの番号を全部（自動採番がぶつからないように）。
           ★取り消した番号も含める＝倉庫の一意制約に where が無いのと同じ扱い★ */
        usedNos: function (docType) {
          var dt = docType || 'invoice';
          return fetchAllQ(function (a, b) {
            return sb.from('pay_invoices').select('no', { count: 'exact' }).eq('doc_type', dt).range(a, b);
          }).then(function (r) {
            if (r && r.error) throw new Error(err(r.error));
            return (r.data || []).map(function (x) { return x.no; });
          });
        },
        /* 下書きの保存（新規も更新も）。★発行済みには使わない（倉庫のトリガが弾く）★ */
        saveDraft: function (inv) {
          return uid().then(function (u) {
            if (!u) return { ok: false, reason: 'ログインしていません' };
            var id = inv.id || newId('iv');
            var row = {
              id: id, account_id: u,
              doc_type: inv.doc_type || 'invoice',
              no: inv.no || '',
              partner_id: inv.partner_id || '',
              issue_ymd: inv.issue_ymd || null,
              due_ymd: inv.due_ymd || null,
              status: 'draft',
              tax_mode: inv.tax_mode, rounding: inv.rounding,
              lines: inv.lines || [], totals: inv.totals || {}, snapshot: inv.snapshot || {},
              data: inv.data || {}, template_id: inv.template_id || '',
              quote_from: inv.quote_from || '',
              deleted_at: null, updated_at: nowIso(),
            };
            return Promise.resolve(sb.from('pay_invoices').upsert(row))
              .then(function (w) { return (w && w.error) ? { ok: false, reason: err(w.error), dup: isDup(w.error) } : { ok: true, id: id }; })
              .catch(function (e) { return { ok: false, reason: err(e), dup: isDup(e) }; });
          });
        },
        /* ★発行★ 番号・合計・写しを決めて固める。
           倉庫が番号の重複で弾いたら bumpNo で1つ進めて出し直す（最後の砦は倉庫）。 */
        issue: function (inv, at) {
          var stamp = at || nowIso();
          return uid().then(function (u) {
            if (!u) return { ok: false, reason: 'ログインしていません' };
            var id = inv.id || newId('iv');
            var no = inv.no || '';
            if (!no) return { ok: false, reason: '請求番号が空です' };
            var base = {
              id: id, account_id: u,
              doc_type: inv.doc_type || 'invoice',
              partner_id: inv.partner_id || '',
              issue_ymd: inv.issue_ymd || null,
              due_ymd: inv.due_ymd || null,
              status: 'issued',
              tax_mode: inv.tax_mode, rounding: inv.rounding,
              lines: inv.lines || [], totals: inv.totals || {}, snapshot: inv.snapshot || {},
              data: inv.data || {}, template_id: inv.template_id || '',
              quote_from: inv.quote_from || '',
              issued_at: stamp, deleted_at: null, updated_at: stamp,
            };
            var tries = 0;
            function attempt(n) {
              base.no = n;
              return Promise.resolve(sb.from('pay_invoices').upsert(base))
                .then(function (w) {
                  if (!(w && w.error)) return { ok: true, id: id, no: n, bumped: tries };
                  if (isDup(w.error) && tries < MAX_BUMP) {
                    var next = DOC.bumpNo(n);
                    // ★数で終わらない番号は自動で進めない（人が決めた番号を勝手にいじらない）
                    if (!next) return { ok: false, reason: 'この番号は既に使われています（' + n + '）。別の番号にしてください', dup: true };
                    tries++;
                    return attempt(next);
                  }
                  return { ok: false, reason: err(w.error), dup: isDup(w.error) };
                })
                .catch(function (e) { return { ok: false, reason: err(e), dup: isDup(e) }; });
            }
            return attempt(no);
          });
        },
        /* 取り消し（★消さない・番号を欠番にしない★） */
        voidIt: function (id, at) {
          var stamp = at || nowIso();
          return Promise.resolve(sb.from('pay_invoices').update({ status: 'void', voided_at: stamp, updated_at: stamp }).eq('id', id))
            .then(function (w) { return (w && w.error) ? { ok: false, reason: err(w.error) } : { ok: true }; })
            .catch(function (e) { return { ok: false, reason: err(e) }; });
        },
        /* 下書きだけ消せる（発行済みは倉庫が2枚で止める） */
        removeDraft: function (id) {
          return Promise.resolve(sb.from('pay_invoices').delete().eq('id', id).eq('status', 'draft'))
            .then(function (w) { return (w && w.error) ? { ok: false, reason: err(w.error) } : { ok: true }; })
            .catch(function (e) { return { ok: false, reason: err(e) }; });
        },
        /* 「送った」記録（発行後も入れられる＝固まる列に入っていない） */
        markSent: function (id, at) {
          var stamp = at || nowIso();
          return Promise.resolve(sb.from('pay_invoices').update({ sent_at: stamp, updated_at: stamp }).eq('id', id))
            .then(function (w) { return (w && w.error) ? { ok: false, reason: err(w.error) } : { ok: true }; })
            .catch(function (e) { return { ok: false, reason: err(e) }; });
        },
      },

      /* ═══ 入金（②では読むだけ。記録の画面は後） ═══
         ★読めなかった時は null を返す＝「0件」と作り分ける（未確認と未入金は違う）★ */
      receipts: {
        list: function () {
          return fetchAllQ(function (a, b) {
            return sb.from('pay_receipts').select('id,invoice_id,invoice_no,ymd,amount,method,memo,deleted_at', { count: 'exact' })
              .is('deleted_at', null).order('ymd', { ascending: true }).range(a, b);
          }).then(function (r) {
            if (r && r.error) return null;      // ★取れなかった＝null（0件と混ぜない）
            return r.data || [];
          }).catch(function () { return null; });
        },
      },

      /* ═══ 共有マスタ（E0 に委譲・ここでは持たない） ═══ */
      org: {
        get: function () { return suite ? suite.org.get() : Promise.resolve(null); },
        save: function (patch) { return suite ? suite.org.save(patch) : Promise.resolve({ ok: false, reason: '共有データ層がありません' }); },
      },
      partners: {
        list: function () { return suite ? suite.partners.list() : Promise.resolve([]); },
        /* ★丸ごと置換しない★ 既存の data（hub が入れた名称・住所など）を残して足す */
        patch: function (id, add) {
          if (!suite) return Promise.resolve({ ok: false, reason: '共有データ層がありません' });
          return suite.partners.list().then(function (list) {
            var cur = list.filter(function (x) { return x.id === id; })[0];
            if (!cur) return { ok: false, reason: 'この取引先が見つかりません' };
            var data = Object.assign({}, cur.data || {}, add || {});
            return suite.partners.upsert({ id: id, sort: cur.sort, data: data });
          });
        },
      },
    };
    return api;
  }

  global.SeikyuStore = { create: create, MAX_BUMP: MAX_BUMP };
})(typeof window !== 'undefined' ? window : globalThis);
