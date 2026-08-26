/* login-gate.test.mjs — ★表の画面(book.html)にも ログインを掛けたか／忘れた人の逃げ道が在るか★
 *
 *  なぜ（2026-08-23 指示役の実測）
 *    ・★book.html は ログインを1本も読んでいなかった＝素通り★
 *      ⇒ ★誰の回数か 数えられない＝人ごとの上限が掛けられない★
 *    ・★「パスワードを忘れた」が 0件★
 *      ⇒ ★忘れた人は 二度と自分のデータに入れない（気づけないまま行き止まり）★
 *      ⇒ 昨日 独自SMTPが入って ★メールが本当に届くようになった＝今日から実際に押される★
 *
 *  ★包み直さない★＝book.html は body の直下に箱が並ぶ作り。丸ごと1つに包むと表の高さが崩れる。
 *  だから ★body に印を付けて CSS で隠す★（隠す1行が 実在するかも ここで数える）。
 *
 *  使い方: node tests/login-gate.test.mjs [--self-test]
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { 注記を外す } from '../scripts/lib/chuki.mjs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
const SELF = process.argv.includes('--self-test');
const OVERRIDE = process.env.EXALLY_GATE_OVERRIDE ? JSON.parse(process.env.EXALLY_GATE_OVERRIDE) : {};
const 読む = (rel) => fs.readFileSync(OVERRIDE[rel] || path.join(ROOT, rel), 'utf8');

let JSDOM;
try { ({ JSDOM } = await import('jsdom')); }
catch { console.log('★jsdomが入っていません。この検証は飛ばせません（SKIPを緑と呼ばない）'); process.exit(1); }

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const AT = async (n, fn) => { try { await fn(); pass++; console.log('  ok   ' + n); } catch (e) { fail++; console.log('  NG   ' + n + '\n       ' + (e && e.message)); } };
const ok = (c, m) => { if (!c) throw new Error(m || 'expected truthy'); };
const eq = (a, b, m) => { if (a !== b) throw new Error((m ? m + ': ' : '') + '期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); };

console.log('');
console.log('[login-gate] ★表の画面にも ログインを掛けたか／忘れた人の逃げ道★');

const book = 読む('book.html');
const loginJs = 読む('js/exally-login.js');
const authJs = 読む('js/auth.js');

/* ── ① 表の画面に 掛かっているか ── */
T('★book.html が ログインの部品を 4本とも読んでいる★', () => {
  for (const 要る of ['supabase-js@2', 'js/supa-config.js', 'js/exally-login.js', 'js/auth.js']) {
    ok(book.indexOf(要る) >= 0, '★' + 要る + ' を読んでいない＝素通り★');
  }
});
T('★ログインが済むまで 中身を隠す1行が 実在する★', () => {
  const 動く所 = 注記を外す(book, { html: true });
  ok(/body\.exally-locked\s*>\s*\*:not\(#loginOv\)\s*\{[^}]*display\s*:\s*none\s*!important/.test(動く所),
    '★隠す1行が無い（ログイン前に 表が見えてしまう）★');
});
T('★auth.js が body に印を付ける（#app が無い画面でも隠せる）★', () => {
  ok(authJs.indexOf('exally-locked') >= 0, '★印を付けていない＝book.html では隠れない★');
  ok(/function\s+lockBody/.test(authJs), 'lockBody が無い');
});

/* ── ② 忘れた人の逃げ道 ── */
T('★「パスワードを忘れた」ボタンが 在る★', () => {
  ok(loginJs.indexOf('パスワードを忘れた') >= 0, '★逃げ道が無い＝二度と入れない★');
  ok(loginJs.indexOf('btnForgot') >= 0, 'ボタンのidが無い');
});
T('★倉庫が resetPasswordForEmail を持つ時だけ 出す（押しても何も起きない物を出さない）★', () => {
  ok(/sb && sb\.auth && sb\.auth\.resetPasswordForEmail[\s\S]{0,200}btnForgot/.test(loginJs),
    '★持っていない時も 出している★');
});
T('★戻ってきた人の見分けが 3通り在る（1通りだと 版が変わった日に袋小路）★', () => {
  for (const [名, 印] of [['自分で付けた目印', 'pwreset=1'], ['Supabaseが付ける', 'type=recovery'], ['Supabaseの合図', 'PASSWORD_RECOVERY']]) {
    ok(loginJs.indexOf(印) >= 0, '★' + 名 + '（' + 印 + '）を 見ていない★');
  }
});
T('★決め終わったら 目印を消す（読み込み直しで また出るのを防ぐ）★', () => {
  ok(/function\s+cleanUrl/.test(loginJs), 'cleanUrl が無い');
  ok(/recoveryOn = false;[\s\S]{0,80}cleanUrl\(\)/.test(loginJs), '★決めた後に 消していない★');
});
T('★送った時は「送った」だけ言う（その住所が登録されているかを 教えない）★', () => {
  const i = loginJs.indexOf('loginResetSent');
  ok(i > 0, '送った画面が無い');
  const 画面 = loginJs.slice(i, i + 700);
  for (const だめ of ['登録されています', '登録されていません', '見つかりません']) {
    ok(画面.indexOf(だめ) < 0, '★' + だめ + ' と言っている（当てられてしまう）★');
  }
});
T('★見た目は Exally のまま（他アプリの名前・色を持ってこない）★', () => {
  for (const よそ of ['Rakually', 'ラクアリー', 'Castally', 'ダイコメ']) {
    ok(loginJs.indexOf(よそ) < 0, '★' + よそ + ' が混ざっている★');
  }
  const i = loginJs.indexOf('loginReset');
  ok(loginJs.slice(i - 400, i + 900).indexOf('Exally') >= 0, '足した画面に Exally の名前が無い');
});

/* ── ③ 本物の画面で 押す ── */
const 画面を作る = (url) => {
  /* ★runScripts を付けないと 部品が window に付かない★（最初これで6本 落ちた） */
  const dom = new JSDOM('<!doctype html><html><head></head><body><div id="x">中身</div></body></html>',
    { url, runScripts: 'dangerously' });
  const w = dom.window;
  const el = w.document.createElement('script');
  el.textContent = loginJs;
  w.document.body.appendChild(el);
  return w;
};
await AT('★実際に載せる：忘れたボタンを押すと 送った画面になる（ログイン画面に戻さない）★', async () => {
  const w = 画面を作る('https://exally.vercel.app/book.html');
  let 送った = null;
  const sb = { auth: {
    resetPasswordForEmail: async (mail, o) => { 送った = { mail, o }; return {}; },
    signInWithPassword: async () => ({ data: {} }),
  } };
  const L = w.ExallyLogin.mount({ app: '表', sb, onLogin() {} });
  const btn = w.document.getElementById('btnForgot');
  ok(btn, '★ボタンが 画面に出ていない★');
  w.document.getElementById('loginEmail').value = 'a@example.com';
  btn.click();
  await new Promise((r) => setTimeout(r, 20));
  ok(送った, '★送っていない★');
  eq(送った.mail, 'a@example.com');
  ok(String(送った.o.redirectTo).indexOf('pwreset=1') >= 0, '★戻り先に 目印が無い★：' + 送った.o.redirectTo);
  ok(String(送った.o.redirectTo).indexOf('exally.vercel.app') >= 0, '★戻り先が 自分のページでない★');
  ok(w.document.getElementById('loginResetSent'), '★送った画面になっていない★');
  ok(!w.document.getElementById('btnLogin'), '★ログイン画面に戻している★');
});
await AT('★目印つきで戻ってきたら いきなり「決める画面」（決めるまで戻さない）★', async () => {
  const w = 画面を作る('https://exally.vercel.app/book.html?pwreset=1');
  let 決めた = null, 入れた = null;
  const sb = { auth: {
    resetPasswordForEmail: async () => ({}),
    updateUser: async (o) => { 決めた = o; return { data: { user: { email: 'a@example.com' } } }; },
  } };
  w.ExallyLogin.mount({ app: '表', sb, onLogin(u) { 入れた = u; } });
  ok(w.document.getElementById('loginReset'), '★決める画面が出ていない（袋小路）★');
  ok(!w.document.getElementById('btnLogin'), 'ログイン画面のままになっている');
  w.document.getElementById('loginNew').value = 'abcdef';
  w.document.getElementById('btnSetPass').click();
  await new Promise((r) => setTimeout(r, 20));
  eq(決めた && 決めた.password, 'abcdef', '★新しいパスワードを 決めていない★');
  ok(入れた, '★決めたのに 入れていない（ログイン画面に戻している）★');
});
await AT('★Supabase の目印(type=recovery)でも 決める画面になる★', async () => {
  const w = 画面を作る('https://exally.vercel.app/book.html#access_token=x&type=recovery');
  w.ExallyLogin.mount({ app: '表', sb: { auth: { resetPasswordForEmail: async () => ({}), updateUser: async () => ({ data: {} }) } }, onLogin() {} });
  ok(w.document.getElementById('loginReset'), '★2通り目の見分けが 効いていない★');
});
await AT('★合図(PASSWORD_RECOVERY)だけでも 決める画面になる★', async () => {
  const w = 画面を作る('https://exally.vercel.app/book.html');
  let 合図 = null;
  w.ExallyLogin.mount({ app: '表', sb: { auth: {
    resetPasswordForEmail: async () => ({}), updateUser: async () => ({ data: {} }),
    onAuthStateChange: (fn) => { 合図 = fn; },
  } }, onLogin() {} });
  ok(合図, '★合図を 受け取る口が 無い★');
  合図('PASSWORD_RECOVERY');
  ok(w.document.getElementById('loginReset'), '★3通り目の見分けが 効いていない★');
});
await AT('★倉庫が古くて resetPasswordForEmail が無い時は ボタンを出さない★', async () => {
  const w = 画面を作る('https://exally.vercel.app/book.html');
  w.ExallyLogin.mount({ app: '表', sb: { auth: { signInWithPassword: async () => ({}) } }, onLogin() {} });
  ok(!w.document.getElementById('btnForgot'), '★押しても何も起きないボタンを 出している★');
  ok(w.document.getElementById('btnLogin'), 'ログイン画面が 出ていない');
});
await AT('★6文字未満は 断る（境界：6文字は通る）★', async () => {
  for (const [pw, 通る] of [['abcde', false], ['abcdef', true]]) {
    const w = 画面を作る('https://exally.vercel.app/book.html?pwreset=1');
    let 決めた = null;
    w.ExallyLogin.mount({ app: '表', sb: { auth: { resetPasswordForEmail: async () => ({}), updateUser: async (o) => { 決めた = o; return { data: {} }; } } }, onLogin() {} });
    w.document.getElementById('loginNew').value = pw;
    w.document.getElementById('btnSetPass').click();
    await new Promise((r) => setTimeout(r, 20));
    eq(!!決めた, 通る, pw + ' の扱いが違う');
  }
});

console.log('');
console.log('  ' + pass + ' 緑 / ' + fail + ' 赤');

if (SELF) {
  const { spawnSync } = await import('node:child_process');
  const os = await import('node:os');
  const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'exally-gate-'));
  console.log('');
  console.log('[self-test] わざと壊して 赤くなるかを数える（★repo は読むだけ★）');
  const BREAKS = [
    ['book.html', '★表の画面から ログインを外す★', (s) => s.replace('<script src="js/auth.js', '<script src="js/auth-none.js')],
    ['book.html', '★隠す1行を消す（ログイン前に 表が見える）★', (s) => s.replace(/body\.exally-locked[^\n]*\n/, '')],
    ['js/auth.js', '★body に印を付けない（book.html では隠れない）★', (s) => s.replace("document.body.classList[on ? 'add' : 'remove']('exally-locked');", '')],
    ['js/exally-login.js', '★忘れたボタンを 消す★', (s) => s.replace('id="btnForgot">パスワードを忘れた', 'id="btnForgotX">忘れない')],
    ['js/exally-login.js', '★倉庫が持っていなくても ボタンを出す★', (s) => s.replace('(sb && sb.auth && sb.auth.resetPasswordForEmail', '(true')],
    ['js/exally-login.js', '★見分けを1通りに減らす（type=recovery を見ない）★', (s) => s.replace('return /(^|[#&?])type=recovery(&|$)/.test(h);', 'return false;')],
    ['js/exally-login.js', '★合図(PASSWORD_RECOVERY)を 受けない★', (s) => s.replace('if (ev === "PASSWORD_RECOVERY") { recoveryOn = true; showResetForm(); }', '')],
    ['js/exally-login.js', '★決めた後に ログイン画面へ戻す（入れたのに戻す）★', (s) => s.replace('      ok((r && r.data && r.data.user) || null);', '      showLoginForm();')],
    ['js/exally-login.js', '★目印を消さない（読み込み直しで また出る）★', (s) => s.replace('      cleanUrl();', '')],
    ['js/exally-login.js', '★戻り先に 目印を付けない★', (s) => s.replace('var back = location.origin + location.pathname + "?" + RESET_MARK;', 'var back = location.origin + location.pathname;')],
    ['js/exally-login.js', '★送った画面で 登録の有無を 教える★', (s) => s.replace('<div class="login-title">パスワードの再設定メールを送りました</div>', '<div class="login-title">登録されています</div>')],
    ['js/exally-login.js', '★5文字でも 通す（境界を外す）★', (s) => s.replace('if (!pw || pw.length < 6)', 'if (!pw)')],
    ['js/exally-login.js', '★他アプリの名前を 持ってくる★', (s) => s.replace('<div class="login-logo">Exally <span>エクサリー</span></div>', '<div class="login-logo">Rakually <span>ラクアリー</span></div>')],
  ];
  let red = 0;
  for (const [rel, name, brk] of BREAKS) {
    const orig = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    const bad = brk(orig);
    if (bad === orig) { console.log('  ★置換できず★  ' + name); continue; }
    const tmpFile = path.join(TMP, rel.replace(/[\\/]/g, '_'));
    fs.writeFileSync(tmpFile, bad, 'utf8');
    const env = Object.assign({}, process.env, { EXALLY_GATE_OVERRIDE: JSON.stringify({ [rel]: tmpFile }) });
    const r = spawnSync(process.execPath, [path.join(__dirname, 'login-gate.test.mjs')], { encoding: 'utf8', env });
    if (r.status !== 0) { red++; console.log('  赤くなった  ' + name); }
    else console.log('  ★素通り★  ' + name);
  }
  for (const rel of ['book.html', 'js/auth.js', 'js/exally-login.js']) {
    const now = fs.readFileSync(path.join(ROOT, rel), 'utf8');
    if (now.includes('auth-none.js') || now.includes('btnForgotX')) {
      console.log('  ★NG★ ' + rel + ' に わざと壊した物が残っている'); process.exit(1);
    }
  }
  try { fs.rmSync(TMP, { recursive: true, force: true }); } catch (e) { /* 消せなくても検査は済んでいる */ }
  console.log('');
  console.log('  ' + red + '/' + BREAKS.length + ' 通りで赤くなった');
  process.exit(red === BREAKS.length ? 0 : 1);
}

process.exit(fail ? 1 : 0);
