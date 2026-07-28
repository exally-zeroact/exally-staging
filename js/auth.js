/* auth.js — ログイン(メール+パスワード/Supabase auth) ＋ 利用権ゲート。hub.js の後に読み込む。
 * Kyually(payslip-app/js/auth.js)と同じ流儀・同じ倉庫・★同じアカウント★＝1ログインで両方使える。
 * Supabase未設定(window.SUPA や supabase-js が無い)なら、ログイン画面を出したまま何もしない
 *   ＝中身を1バイトも出さない(未ログインでデータ画面を見せない)。
 */
(function (global) {
  'use strict';
  var hasSupa = !!(global.SUPA && global.SUPA.url && global.SUPA.key && global.supabase);

  var st = document.createElement('style');
  st.textContent = '#auth-overlay{position:fixed;inset:0;z-index:1000;background:linear-gradient(160deg,#EAF6EF,#F0FAF4);display:none;align-items:center;justify-content:center;padding:20px;font-family:"Noto Sans JP",sans-serif;}'
    + '.auth-card{width:100%;max-width:360px;background:#fff;border:1px solid #d4eae0;border-radius:18px;box-shadow:0 10px 36px rgba(30,60,40,.12);padding:26px 22px;}'
    + '.auth-logo{font-family:\'DM Mono\',monospace;font-size:30px;font-weight:500;letter-spacing:-0.5px;color:#52B788;text-align:center;margin-bottom:10px;}'
    + '.auth-card p.lead{font-size:12px;color:#7aa08c;text-align:center;margin:8px 0 16px;line-height:1.7;}'
    + '.auth-card input{width:100%;padding:12px 12px;margin-bottom:10px;border:1.5px solid #d4eae0;border-radius:10px;font-size:16px;color:#1a4a2e;-webkit-appearance:none;}'
    + '.auth-card input:focus{outline:none;border-color:#52B788;}'
    + '#auth-msg{font-size:12px;min-height:16px;margin:2px 0 10px;text-align:center;}'
    + '.auth-card .b1{width:100%;padding:12px;border:none;border-radius:10px;background:#3D9E72;color:#fff;font-size:15px;font-weight:700;cursor:pointer;}'
    + '.auth-card .b2{width:100%;padding:11px;margin-top:8px;border:1.5px solid #d4eae0;border-radius:10px;background:#fff;color:#3D6B53;font-size:14px;font-weight:700;cursor:pointer;}';
  document.head.appendChild(st);

  var ov = document.createElement('div'); ov.id = 'auth-overlay';
  ov.innerHTML = '<div class="auth-card"><div class="auth-logo">Exally</div>'
    + '<p class="lead">ログインすると、どの端末でも同じ内容で使えます。<br>給料明細アプリと同じメール・パスワードです。</p>'
    + '<input id="auth-email" type="email" placeholder="メールアドレス" autocomplete="username">'
    + '<input id="auth-pw" type="password" placeholder="パスワード（6文字以上）" autocomplete="current-password">'
    + '<div id="auth-msg"></div>'
    + '<button class="b1" id="auth-login" type="button">ログイン</button>'
    + '<button class="b2" id="auth-signup" type="button">新規登録（はじめての方）</button></div>';
  document.body.appendChild(ov);

  function $(id) { return document.getElementById(id); }
  // ★中身(.app)はログインが済むまで hidden のまま＝未ログインで画面を見せない
  function show() { ov.style.display = 'flex'; var a = $('app'); if (a) a.hidden = true; }
  function hide() { ov.style.display = 'none'; var a = $('app'); if (a) a.hidden = false; }
  function msg(t, err) { var m = $('auth-msg'); if (!m) return; m.textContent = t || ''; m.style.color = err ? '#C0392B' : '#3D6B53'; }
  function jpErr(s) {
    s = String(s || '');
    if (/Invalid login/i.test(s)) return 'メールかパスワードが違います';
    if (/already registered|User already/i.test(s)) return 'このメールは登録済みです。ログインしてください';
    if (/at least 6/i.test(s)) return 'パスワードは6文字以上にしてください';
    if (/valid email/i.test(s)) return 'メールアドレスの形式が正しくありません';
    if (/Failed to fetch|NetworkError/i.test(s)) return 'ネットに繋がりませんでした';
    return s;
  }

  if (!hasSupa) { show(); msg('この端末では接続設定が読み込めませんでした', true); return; }

  var sb = global.supabase.createClient(global.SUPA.url, global.SUPA.key);
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
    ov.innerHTML = '<div class="auth-card"><div class="auth-logo">Exally</div>'
      + '<p class="lead" style="color:#92500A;font-weight:700;margin:8px 0 14px">' + m.title + '</p>'
      + (m.body ? '<p class="lead" style="margin-top:-8px">' + m.body + '</p>' : '')
      + '<button class="b2" id="auth-lock-out" type="button">別のアカウントでログイン</button></div>';
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

  $('auth-login').onclick = function () {
    var e = $('auth-email').value.trim(), p = $('auth-pw').value;
    if (!e || !p) { msg('メールとパスワードを入力してください', true); return; }
    msg('ログイン中...');
    sb.auth.signInWithPassword({ email: e, password: p }).then(function (r) {
      if (r.error) msg(jpErr(r.error.message), true); else afterLogin(e);
    });
  };
  $('auth-signup').onclick = function () {
    var e = $('auth-email').value.trim(), p = $('auth-pw').value;
    if (!e) { msg('メールアドレスを入力してください', true); return; }
    if (p.length < 6) { msg('パスワードは6文字以上にしてください', true); return; }
    msg('登録中...');
    sb.auth.signUp({ email: e, password: p }).then(function (r) {
      if (r.error) msg(jpErr(r.error.message), true);
      else if (!r.data || !r.data.session) msg('確認メールを送りました。メールのリンクを開いてからログインしてください');
      else afterLogin(e);
    });
  };
  $('auth-pw').addEventListener('keydown', function (ev) { if (ev.key === 'Enter') $('auth-login').click(); });

  // 起動時: セッションがあればそのまま、無ければログイン画面
  sb.auth.getSession().then(function (r) {
    var s = r && r.data && r.data.session;
    if (s) afterLogin((s.user && s.user.email) || ''); else show();
  }).catch(function () { show(); });
  sb.auth.onAuthStateChange(function (_e, s) { if (!s) show(); });

})(window);
