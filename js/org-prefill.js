/* org-prefill.js — 共有の自社情報(pay_org)を、請求書/見積の発行者欄に自動で入れる。
 *
 * ねらい（Exallyの旗＝二度手間ゼロ）:
 *   「共有データ ▸ 会社」で1回入れた 屋号・住所・電話・インボイス番号 を、
 *   請求書や見積で毎回書き直させない。
 *
 * ★守る事★
 *   ・pay_org が未設定なら何もしない（空欄のまま。勝手に埋めない＝捏造しない）
 *   ・すでに人が入力している欄は上書きしない（手入力が勝つ）
 *   ・入れた後も自由に直せる（ただの初期値）
 *   ・法定データ(消費税率など statutory)の既存の読みには一切関与しない
 *   ・未ログイン/通信失敗でも請求書は今まで通り使える（黙って何もしない）
 *
 * 【利用】ブラウザ: <script src="js/org-prefill.js"> を supa-config/supabase-js の後に置くだけ
 *        Node(テスト): require/import で pickOrgFields / planPrefill を使う
 */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.OrgPrefill = api;
  // ブラウザなら自動で走る（読み込むだけで効く）
  if (typeof window !== 'undefined' && typeof document !== 'undefined') api.auto();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // pay_org の data → 請求書/見積の入力欄id への対応
  var MAP = {
    yago: 'issuerName',
    addr: 'issuerAddr',
    tel: 'issuerTel',
    invoiceNo: 'invoiceNo'
  };

  function trim(v) { return String(v == null ? '' : v).trim(); }

  // pay_org の data から、請求書で使う項目だけ取り出す（空/空白は「未設定」）
  function pickOrgFields(orgData) {
    var out = {};
    if (!orgData || typeof orgData !== 'object') return out;
    Object.keys(MAP).forEach(function (k) {
      var v = trim(orgData[k]);
      if (v) out[MAP[k]] = v;
    });
    return out;
  }

  // 「どの欄に何を入れるか」を決める。current = 今の画面の値 {id: value}
  //   ・画面に無い欄は触らない
  //   ・すでに入力がある欄は上書きしない
  function planPrefill(fields, current) {
    var plan = {};
    fields = fields || {}; current = current || {};
    Object.keys(fields).forEach(function (id) {
      if (!(id in current)) return;          // その欄が画面に無い
      if (trim(current[id])) return;         // 人が入れている＝尊重する
      plan[id] = fields[id];
    });
    return plan;
  }

  /* ── ブラウザ側 ── */

  function readCurrent(ids) {
    var cur = {};
    ids.forEach(function (id) {
      var el = document.getElementById(id);
      if (el) cur[id] = el.value;
    });
    return cur;
  }

  function apply(plan) {
    var n = 0;
    Object.keys(plan).forEach(function (id) {
      var el = document.getElementById(id);
      if (!el) return;
      el.value = plan[id];
      // 画面側の再計算/保存フックに気づかせる（プレビューや自動保存が動くように）
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      n++;
    });
    return n;
  }

  // pay_org を読む（未ログイン・未設定・通信失敗はすべて「何もしない」）
  function fetchOrg() {
    try {
      var w = window;
      if (!(w.SUPA && w.SUPA.url && w.SUPA.key && w.supabase)) return Promise.resolve(null);
      var sb = w.supabase.createClient(w.SUPA.url, w.SUPA.key);
      return sb.auth.getUser().then(function (r) {
        var uid = r && r.data && r.data.user && r.data.user.id;
        if (!uid) return null;                                  // 未ログイン＝何もしない
        return sb.from('pay_org').select('data').maybeSingle().then(function (q) {
          if (q && q.error) return null;
          return (q && q.data && q.data.data) || null;
        });
      }).catch(function () { return null; });
    } catch (e) { return Promise.resolve(null); }
  }

  // 読み込むだけで効く。発行者欄がある画面(請求書/見積)でのみ動く。
  function auto() {
    function run() {
      var ids = Object.keys(MAP).map(function (k) { return MAP[k]; });
      var exists = ids.some(function (id) { return !!document.getElementById(id); });
      if (!exists) return;                                       // 発行者欄が無い画面＝対象外
      fetchOrg().then(function (org) {
        if (!org) return;                                        // 未設定/未ログイン＝空のまま
        var plan = planPrefill(pickOrgFields(org), readCurrent(ids));
        var n = apply(plan);
        if (n) notify(n);
      });
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', run);
    else run();
  }

  // 「勝手に入った」と驚かせないよう、入れたことを小さく知らせる（直せることも書く）
  function notify(n) {
    try {
      var el = document.createElement('div');
      el.textContent = '共有データの会社情報を ' + n + '箇所に入れました（直せます）';
      el.style.cssText = 'position:fixed;left:50%;transform:translateX(-50%);bottom:76px;z-index:60;'
        + 'background:#1A4A2E;color:#fff;font-size:12px;line-height:1.6;padding:10px 16px;border-radius:12px;'
        + 'max-width:88%;box-shadow:0 6px 20px rgba(30,60,40,.25);font-family:"Noto Sans JP",sans-serif;';
      document.body.appendChild(el);
      setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 4000);
    } catch (e) {}
  }

  return { pickOrgFields: pickOrgFields, planPrefill: planPrefill, auto: auto, MAP: MAP };
});
