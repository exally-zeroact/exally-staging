/* statutory-hydrate.js — Exally: 中央Supabase statutory の法定値を各libに注入(payslip-appと同方式)。
 *   ・全アプリ共通の単一ソース(倉庫=Exallyプロジェクト)。読取専用(anon)・書込みは中央のみ。
 *   ・取れない/オフライン/不正なら各libのハードコードで動く(フォールバック)。
 *   ・Exallyは令和7基準(各libのNENDO)なので 社保/雇用/最賃=2025・割増=2023・消費税=2019 の行で注入(エラは跳ばさない=数字は誤→正の是正のみ)。
 *   ・置き場所: 各HTMLで統計lib群の後・アプリ本体scriptの前に読み込む。
 */
(function () {
  'use strict';
  var SUPA = {
    url: 'https://tnfwipbgfgjaymlszeid.supabase.co',
    key: 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InRuZndpcGJnZmdqYXltbHN6ZWlkIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODE1Nzk4MzQsImV4cCI6MjA5NzE1NTgzNH0.zhKPLSlW4zxsdjsXNvqDHvtP3wBqp-EKaxbjqLGW_ek'
  };
  // Exally各libが対象とするエラ(年度)。この年度の行だけ各libへ注入する。
  var ERA = { saitei_chingin: 2025, shakaihoken: 2025, koyo: 2025, warimashi: 2023, shouhizei: 2019 };
  function lib(name) { try { return (typeof window !== 'undefined' && window[name]) || (new Function('try{return ' + name + '}catch(e){return null}'))(); } catch (e) { return null; } }

  function distribute(rows) {
    if (!rows || !rows.length) return;
    var SAI = lib('SAITEI_CHINGIN'), SHH = lib('SHAKAIHOKEN_HYO'), KOY = lib('KOYOHOKEN_RITSU'), ROU = lib('ROUKI_RITSU'), SHO = lib('SHOUHIZEI_RITSU');
    rows.forEach(function (r) {
      try {
        if (r.kind === 'saitei_chingin' && r.year === ERA.saitei_chingin && SAI && SAI.hydrate) SAI.hydrate(r.data);
        else if (r.kind === 'shakaihoken' && r.year === ERA.shakaihoken && SHH && SHH.hydrate) SHH.hydrate(r.data);
        else if (r.kind === 'koyo' && r.year === ERA.koyo && KOY && KOY.hydrate) KOY.hydrate(r.data);
        else if (r.kind === 'warimashi' && r.year === ERA.warimashi && ROU && ROU.hydrate) ROU.hydrate(r.data);
        else if (r.kind === 'shouhizei' && r.year === ERA.shouhizei && SHO && SHO.hydrate) SHO.hydrate(r.data);
      } catch (e) { /* 1件失敗しても他はフォールバックで継続 */ }
    });
    if (typeof window.onStatutoryHydrated === 'function') { try { window.onStatutoryHydrated(); } catch (e) { } } // 再計算フック(任意)
  }

  function boot(sb) {
    try {
      sb.from('statutory').select('kind,year,data')
        .then(function (res) { if (res && res.data) distribute(res.data); })
        .catch(function () { /* フォールバック */ });
    } catch (e) { /* フォールバック */ }
  }
  function tryReady() {
    if (typeof window === 'undefined' || !(window.supabase && window.supabase.createClient)) return false;
    try { boot(window.supabase.createClient(SUPA.url, SUPA.key)); return true; } catch (e) { return false; }
  }
  if (!tryReady()) {
    // supabase-js 未ロードならCDNから読み込んでから実行(失敗時はフォールバックのまま)
    try {
      var s = document.createElement('script');
      s.src = 'https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2';
      s.onload = tryReady; s.onerror = function () { };
      document.head.appendChild(s);
    } catch (e) { /* フォールバック */ }
  }
})();
