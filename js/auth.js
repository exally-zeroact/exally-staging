/* auth.js — ログイン(メール+パスワード/Supabase auth) ＋ 利用権ゲート。hub.js の後に読み込む。
 * Kyually(payslip-app/js/auth.js)と同じ流儀・同じ倉庫・★同じアカウント★＝1ログインで両方使える。
 * Supabase未設定(window.SUPA や supabase-js が無い)なら、ログイン画面を出したまま何もしない
 *   ＝中身を1バイトも出さない(未ログインでデータ画面を見せない)。
 */
(function (global) {
  'use strict';
  var hasSupa = !!(global.SUPA && global.SUPA.url && global.SUPA.key && global.supabase);

  // ログイン画面は全アプリ共通の部品(js/exally-login.js)。見た目も文言もそこが一次情報。
  var LOGIN = null;
  var ov = null;
  function mountLogin(sbForLogin) {
    if (LOGIN) return LOGIN;
    LOGIN = global.ExallyLogin.mount({
      app: 'ホーム',
      sb: sbForLogin,
      note: '売上管理・代行請求・給料明細も、同じメールとパスワードで入れます。',
      onLogin: function (user) { afterLogin((user && user.email) || ''); }
    });
    ov = LOGIN.el;
    return LOGIN;
  }

  function $(id) { return document.getElementById(id); }
  // ★中身(.app)はログインが済むまで hidden のまま＝未ログインで画面を見せない
  /* ★#app を持たない画面（book.html＝表そのもの）でも 中身を見せない★（2026-08-23）
     book.html は body の直下に25個の箱が並ぶ作りで、★丸ごと1つに包み直すと 表の高さの計算が崩れる★。
     だから ★body に印を付けて CSS で隠す★（包み直さない）。
     隠す1行は その画面の <style> に置く（[hidden] は class の display に負けるため）。 */
  function lockBody(on) {
    try { document.body.classList[on ? 'add' : 'remove']('exally-locked'); } catch (e) { /* 画面が無い時は何もしない */ }
  }
  function show() { if (LOGIN) LOGIN.show(); var a = $('app'); if (a) a.hidden = true; lockBody(true); }
  function hide() { if (LOGIN) LOGIN.hide(); var a = $('app'); if (a) a.hidden = false; lockBody(false); }
  function msg(t, err) { if (err && LOGIN) LOGIN.error(t || ''); }
  function jpErr(s) {
    if (global.ExallyLogin) return global.ExallyLogin.friendly({ message: s });
    s = String(s || '');
    if (/Invalid login/i.test(s)) return 'メールかパスワードが違います';
    if (/already registered|User already/i.test(s)) return 'このメールは登録済みです。ログインしてください';
    if (/at least 6/i.test(s)) return 'パスワードは6文字以上にしてください';
    if (/valid email/i.test(s)) return 'メールアドレスの形式が正しくありません';
    if (/Failed to fetch|NetworkError/i.test(s)) return 'ネットに繋がりませんでした';
    return s;
  }

  var NG = function () {
    return Promise.resolve({ error: { message: '接続設定が読み込めませんでした' } });
  };
  if (!hasSupa) {
    mountLogin({ auth: { signInWithPassword: NG, signUp: NG } });
    show();
    msg('この端末では接続設定が読み込めませんでした', true);
    return;
  }

  var sb = global.supabase.createClient(global.SUPA.url, global.SUPA.key);
  /* ★4 事故止め（2026-08-25）★ AIの回数は ★人ごと★に数える。
     ★誰かを AIの窓口へ伝えるのに 入口が1本 要る★（同じ回線の人を まとめて1人にしないため）。
     ★出すのは この入れ物だけ★＝鍵や メールを 画面に置かない。 */
  global.Auth = global.Auth || {};
  global.Auth.sb = sb;
  mountLogin(sb);
  var APP = 'suite';
  var curEmail = '';

  // 利用権ゲート: 停止なら使わせない。取れなかった時は"締めない"(誤ロックで締め出さない)
  function gateCheck(sd) {
    if (!(sd && global.Access)) return Promise.resolve({ ok: true, reason: 'nogate' });
    return sd.entitlements.get(APP).then(function (acc) {
      if (!acc) return sd.entitlements.ensure(APP).then(function () { return { ok: true, reason: 'new' }; });
      return global.Access.accessState(acc);
    }).catch(function () { return { ok: true, reason: 'error' }; });
  }

  function showLock() {
    var m = (global.Access && global.Access.lockMessage) ? global.Access.lockMessage() : { title: 'このアカウントは現在ご利用いただけません', body: '' };
    ov.innerHTML = '<div class="login-card"><div class="login-logo">Exally <span>エクサリー</span></div>'
      + '<div class="login-mid" style="color:#92500A;font-weight:700">' + m.title + '</div>'
      + (m.body ? '<div class="login-note">' + m.body + '</div>' : '')
      + '<button class="login-btn login-btn-sub" style="margin-top:14px" id="auth-lock-out" type="button">別のアカウントでログイン</button></div>';
    show();
    var lo = $('auth-lock-out'); if (lo) lo.onclick = function () { sb.auth.signOut().then(function () { location.reload(); }); };
  }

  function showHeader() {
    var h = $('hd-user'); if (!h) return;
    // メールは1行で省略表示(誰で入っているか分かればよい)＋ログアウトは常に見える位置に
    h.innerHTML = '<span class="hd-mail" title="' + (curEmail || '') + '">' + (curEmail || '') + '</span>'
      + '<span class="hd-out" id="auth-logout">ログアウト</span>';
    var lo = $('auth-logout');
    if (lo) lo.onclick = function () { sb.auth.signOut().then(function () { location.reload(); }); };
  }

  function afterLogin(email) {
    curEmail = email || curEmail;
    if (!(global.Hub && global.Hub.attach)) { hide(); return; }
    global.Hub.attach(sb).then(function (sd) {
      return gateCheck(sd).then(function (gate) {
        if (gate && !gate.ok) { showLock(); return; }   // 停止アカウント=中身を触らせない
        hide();
        showHeader();
      });
    }).catch(function (e) { msg(jpErr(e && e.message), true); });
  }

  // 起動時: セッションがあればそのまま、無ければログイン画面
  sb.auth.getSession().then(function (r) {
    var s = r && r.data && r.data.session;
    if (s) afterLogin((s.user && s.user.email) || ''); else show();
  }).catch(function () { show(); });
  sb.auth.onAuthStateChange(function (_e, s) { if (!s) show(); });

})(window);
