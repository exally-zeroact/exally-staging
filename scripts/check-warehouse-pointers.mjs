import { 注記を外す } from './lib/chuki.mjs';
/* check-warehouse-pointers.mjs — ★倉庫(Supabase)の向き先を、6か所×全アプリで数える★
 *
 * なぜ必要か（2026-08-07・指示役）:
 *   「テストrepoが本番倉庫を指していないか」を①アプリのコードだけ見て「事故0」と報告した。
 *   ★実際に壊れていたのは③GitHub Actions の金庫★だった（アマかせのバッチが旧倉庫に書き続けていた）。
 *   向き先は1か所ではなく★6か所★にある。1か所だけ見て安心すると、必ず残りで事故る。
 *
 *     ① アプリのコード（js/supa-config.js・HTML直書き）
 *     ② Vercel の環境変数（配信側）
 *     ③ GitHub Actions の金庫（バッチ）        ← ここが壊れていた
 *     ④ 認証の戻り先 許可リスト
 *     ⑤ Edge Function
 *     ⑥ 見張りの道具そのもの
 *
 * ★測れなかったマスは 🟡未測定 と出す。0件・異常なしにしない。★
 *   （「鍵が無いから測れなかった」を緑と呼ぶと、見張りが嘘をつく）
 *
 * ★全部 読むだけ。他のアプリのリポジトリにも倉庫にも、1文字も書かない。★
 *   使うのは GET だけ（配信物の取得 / GitHubの読み取り / Supabaseの読み取り）。
 *
 * 鍵（無くても動く。無い所は 🟡 になる）:
 *   GITHUB_TOKEN          … 公開repoのファイルを読む（Actionsでは自動で入る）
 *   GH_SECRETS_TOKEN      … ③の登録日を読む（repo管理権限のPAT。無ければ🟡）
 *   VERCEL_TOKEN          … ②の環境変数の名前と更新日（無ければ🟡）
 *   SUPABASE_ACCESS_TOKEN … ⑤の一覧と版（無ければ、関数URLを叩く簡易判定にする）
 *
 * 使い方: node scripts/check-warehouse-pointers.mjs
 *         node scripts/check-warehouse-pointers.mjs --json
 *         node scripts/check-warehouse-pointers.mjs --self-test  ★判定が空振りしていないか
 *
 * 終了コード: 🔴があれば 3 ／ それ以外は 0（🟡の本数は必ず最後に出す）
 */

/* ══════════ 何を正とするか（ここだけが「答え」） ══════════ */
export const PROD_REF = 'tnfwipbgfgjaymlszeid';
export const TEST_REF = 'khawdrnvssdenumbiwfg';

/* アプリ × 本番/テスト。★url は「向き先が書いてある物」を直接指す★ */
const APPS = [
  { app: 'Exally',        env: '本番',   host: 'https://exally.vercel.app',                        cfg: '/js/supa-config.js',  want: PROD_REF },
  { app: 'Exally',        env: 'テスト', host: 'https://exally-zeroact.github.io/exally-staging',  cfg: '/js/supa-config.js',  want: TEST_REF },
  // ★2026-08-07 Vercel版の staging は畳んだ（司さんOK）★
  //   git連携が無く手打ちでしか更新されない＝黙って古くなる形だった。テストの配信は github.io の1本。
  //   ここに残すと「消えた住所」を毎週叩いて 🟡 が鳴り続けるので、行ごと外す。
  { app: '給与 kyuyo',    env: '本番',   host: 'https://exally.vercel.app',                        cfg: '/js/supa-config.js',  want: PROD_REF, note: 'Exallyと同居（同じ設定ファイル）' },
  { app: '給与 kyuyo',    env: 'テスト', host: 'https://exally-zeroact.github.io/exally-staging',  cfg: '/js/supa-config.js',  want: TEST_REF, note: 'Exallyと同居（同じ設定ファイル）' },
  { app: '代行請求',      env: '本番',   host: 'https://daikou-seikyu.vercel.app',                 cfg: '/daikou-seikyu.html', want: PROD_REF, owner: 'ダイコメ' },
  { app: '代行請求',      env: 'テスト', host: null, want: null, owner: 'ダイコメ', note: '★テスト環境は無い（ダイコメの製品・Exallyは触らない）' },
  { app: 'ダイコメ',      env: '本番',   host: 'https://daikou-app.vercel.app',                    cfg: '/js/dk-config.js',    want: PROD_REF, owner: 'ダイコメ' },
  { app: 'ダイコメ',      env: 'テスト', host: 'https://daikou-app-test.vercel.app',               cfg: '/js/dk-config.js',    want: TEST_REF, owner: 'ダイコメ' },
  { app: 'ダイコメ事務所', env: '本番',  host: 'https://daikome-jimusho.vercel.app',               cfg: '/js/dk-config.js',    want: PROD_REF, owner: 'ダイコメ' },
  { app: 'ダイコメ事務所', env: 'テスト', host: 'https://daikome-jimusho-test.vercel.app',         cfg: '/js/dk-config.js',    want: TEST_REF, owner: 'ダイコメ' },
  { app: '飲み屋',        env: '本番',   host: 'https://nomiya-app.vercel.app',                    cfg: '/js/supa-config.js',  want: PROD_REF, owner: '飲み屋' },
  { app: '飲み屋',        env: 'テスト', host: 'https://nomiya-app-test.vercel.app',               cfg: '/js/supa-config.js',  want: TEST_REF, owner: '飲み屋' },
  // ★アマかせは repo に住所を持たない。Vercelの環境変数が /api/config に出てくる＝ここが①と②を兼ねる
  { app: 'アマかせ',      env: '本番',   host: 'https://amazon-ads-automation-lyart.vercel.app',   cfg: '/api/config',         want: PROD_REF, owner: 'アマかせ', viaEnv: true },
  { app: 'アマかせ',      env: 'テスト', host: 'https://amazon-ads-automation-test.vercel.app',    cfg: '/api/config',         want: TEST_REF, owner: 'アマかせ', viaEnv: true },
];

/* ③ バッチ。repo と、そのrepoが向くべき倉庫 */
const BATCH_REPOS = [
  { repo: 'exally-zeroact/exally',                     want: PROD_REF, env: '本番' },
  { repo: 'exally-zeroact/exally-staging',             want: TEST_REF, env: 'テスト' },
  { repo: 'exally-zeroact/daikou-seikyu',              want: PROD_REF, env: '本番', owner: 'ダイコメ' },
  { repo: 'exally-zeroact/Daikou-app',                 want: PROD_REF, env: '本番', owner: 'ダイコメ' },
  { repo: 'exally-zeroact/Daikou-app-test',            want: TEST_REF, env: 'テスト', owner: 'ダイコメ' },
  { repo: 'exally-zeroact/nomiya-app',                 want: PROD_REF, env: '本番', owner: '飲み屋' },
  { repo: 'exally-zeroact/nomiya-app-test',            want: TEST_REF, env: 'テスト', owner: '飲み屋' },
  { repo: 'exally-zeroact/amazon-ads-automation',      want: PROD_REF, env: '本番', owner: 'アマかせ', private: true },
  { repo: 'exally-zeroact/amazon-ads-automation-test', want: TEST_REF, env: 'テスト', owner: 'アマかせ', private: true },
];

/* ★引っ越しの日★ これより前に登録された Supabase系の鍵は「引っ越し前」の疑い */
const MIGRATION_DAY = '2026-08-06';

/* ⑤ 配信されている Edge Function（本番・テストとも同じ4本のはず） */
const EDGE_FUNCS = ['dk-issue-license', 'dk-register-company', 'dk-sync-jobs', 'dk-customers'];

/* ④ 戻り先を確かめる住所
   ★2026-08-24 に決まりを変えた（指示役）★
     前 … 「★全アプリのURLが 両方の倉庫に入っている事★」＝混ざりを永久に見逃す形だった。
           環境の混ざりを外した日に ★9件が赤★になり、しかも ★客への害は0★ だった。
     今 … ①★そのアプリが向いている倉庫に 自分のURLが在るか★（無ければ 忘れた人が入れない）
           ②★向いていない倉庫に 自分のURLが居ないか★（居たら ★混ざり★＝別の環境へ飛べてしまう）
   ★向き先は 実物のコードから読む（cfg）。表に手で書かない★ */
const REDIRECT_HOSTS = [
  { app: 'Exally',         url: 'https://exally.vercel.app/hub.html',
    cfg: 'https://exally.vercel.app/js/supa-config.js' },
  { app: '給与 kyuyo',     url: 'https://exally.vercel.app/kyuyo/admin.html',
    cfg: 'https://exally.vercel.app/js/supa-config.js' },
  { app: 'Exally テスト',  url: 'https://exally-zeroact.github.io/exally-staging/hub.html',
    cfg: 'https://exally-zeroact.github.io/exally-staging/js/supa-config.js' },
  { app: '代行請求',       url: 'https://daikou-seikyu.vercel.app/daikou-seikyu.html',
    cfg: 'https://daikou-seikyu.vercel.app/daikou-seikyu.html' },   /* ★向き先は この画面に直書き★ */
  { app: 'ダイコメ',       url: 'https://daikou-app.vercel.app/',
    cfg: 'https://daikou-app.vercel.app/js/dk-config.js' },
  { app: 'ダイコメ事務所', url: 'https://daikome-jimusho.vercel.app/login.html',
    cfg: 'https://daikome-jimusho.vercel.app/js/dk-config.js' },
  { app: 'アマかせ',       url: 'https://amazon-ads-automation-lyart.vercel.app/',
    cfg: 'https://amazon-ads-automation-lyart.vercel.app/api/config' },   /* ★配信側の環境変数を返す口★ */
  { app: 'アマかせ テスト', url: 'https://amazon-ads-automation-test.vercel.app/',
    cfg: 'https://amazon-ads-automation-test.vercel.app/api/config' },
];


/* ══════════ 判定の中身（純関数・self-testで作り物を通せる） ══════════ */

/** 文章から Supabase の向き先(ref)を取り出す。形は3通りある。 */
export function refsIn(text) {
  const out = new Set();
  let m;
  const url = /https:\/\/([a-z0-9]{16,})\.supabase\.co/g;
  while ((m = url.exec(text))) out.add(m[1]);
  const key = /"ref"\s*:\s*"([a-z0-9]{16,})"/g;
  while ((m = key.exec(text))) out.add(m[1]);
  const jwt = /eyJ[A-Za-z0-9_-]+\.([A-Za-z0-9_-]{20,})\./g;
  while ((m = jwt.exec(text))) {
    try {
      const b = m[1].replace(/-/g, '+').replace(/_/g, '/');
      const o = JSON.parse(Buffer.from(b + '='.repeat((4 - (b.length % 4)) % 4), 'base64').toString('utf8'));
      if (o && typeof o.ref === 'string') out.add(o.ref);
    } catch (_) { /* 読めない物は無視 */ }
  }
  return [...out];
}

/** ★裸の文字列（const PROD_REF = '…' や、実測メモのコメント）★
 *  これも見ないと `auth-mail-otp.mjs` の本番直書き（PROJECT = '…'）を見落とす。
 *  ただし ★「書いてある＝そこへ繋いでいる」ではない★:
 *    ・`nomiya/tests/nomiya-deploy.test.js` は本番refを「ここでは絶対に走らせない」の見張りとして持つ（正しい）
 *    ・`auth-redirect-allow.mjs` は本番とテストの両方を扱うのが正しい形（8/7に直したばかり）
 *    ・`verify-statutory.mjs` の khaw… は実測メモのコメント
 *  だから ★裸の文字列は 🔴 にしない。🟡（理由を書くまで未確認）にする★。
 *  URL・鍵の形（＝実際に繋ぐ形）だけが 🔴。 */
export function bareRefsIn(text) {
  return [PROD_REF, TEST_REF].filter((r) => new RegExp('\\b' + r + '\\b').test(text));
}

/** ★★その 住所が「何の 役目で」書いてあるか★★（2026-09-07・指示役の 裁定）
 *
 *  ★前は「在る／無い」で 分けていた★
 *    ・繋ぐ形（URL・鍵）    … 🔴
 *    ・裸の 文字列          … 🟡
 *  ⇒★★これだと 2つとも 外す★★
 *    ①`js/dk-config.js` の 住所は ★ただの 定数★（fetch の 引数では ない）
 *      ⇒★一番 大事な 1本が すり抜ける★
 *    ②`tests/…` の `not.toContain('…')` は ★反対側を 指していない事を 見る 試験★
 *      ⇒★仲間を 敵と 数える★（2026-09-07 に 実際に 誤報した）
 *  ⇒★★『在る／無い』では なく『何の 役目で 在るか』で 分ける★★
 *
 *  🔴 … ★設定の ファイル★の 中の 住所／`fetch`・`createClient`・`request` の 引数
 *  🟡 … ★試験の 中の 期待値★（`not.toContain(...)`／`toBe(...)`／名前→住所 の 対応表）
 *  ★どちらでも「なぜ そう 数えたか」を 1行 出す★
 */
export function 設定のファイルか(path) {
  const p = String(path);
  /* ★画面（.html）の 中に 直に 書いて ある 住所も「設定」★（2026-09-07 指示役）
     ⇒ 代行請求は `daikou-seikyu.html` の 中に 住所を 直書きしている
     ⇒★今は 本番も テスト線も 正しい★／★でも 明日 反対に なっても 誰も 気づかない★
     ⇒★『起きていない』と『起きない』は 別★ */
  if (/\.html?$/i.test(p)) return true;
  return /(^|\/)[a-z0-9_-]*config[a-z0-9_-]*\.(m?js|ts|json)$/i.test(p);
}
export function 繋ぐ呼び出しの中か(text, ref) {
  const 行 = String(text).split(/\r?\n/);
  for (let i = 0; i < 行.length; i++) {
    if (行[i].indexOf(ref) < 0) continue;
    /* ★呼び出しは 前の 行に またがる★ので 2行 前まで 見る */
    const 窓 = 行.slice(Math.max(0, i - 2), i + 1).join(' ');
    if (/(fetch|createClient|request|axios|got|open)\s*\(/.test(窓)) return true;
  }
  return false;
}
export function 試験の期待値か(path, text, ref) {
  if (!/(^|\/)(tests?|__tests__|spec)\//.test(String(path))) return false;
  const 行 = String(text).split(/\r?\n/);
  for (let i = 0; i < 行.length; i++) {
    if (行[i].indexOf(ref) < 0) continue;
    const 窓 = 行.slice(Math.max(0, i - 2), i + 1).join(' ');
    /* 期待値の 書き方／名前→住所 の 対応表 */
    if (/(expect|assert|to(Contain|Be|Equal|Match|StrictEqual)|not\.)/.test(窓)) continue;
    if (/['\"][^'\"]+['\"]\s*:\s*['\"][^'\"]*/.test(行[i])) continue;
    return false;                                   /* 期待値では ない 行が 1つでも 在れば false */
  }
  return true;
}
/** @returns {{mark:'🔴'|'🟡', 訳:string}} */
export function 役目で分ける(path, text, ref) {
  /* ★★注記（コメント）は 先に 外す★★（2026-09-07 実測で 踏んだ）
     ⇒ 飲み屋の `js/supa-config.js` は 頭の 注記に
       「テスト用DB(khaw…)とは 別の…」と ★訳を 書いている★
     ⇒ 注記を 外さずに 数えると ★正しい 設定を 🔴 に する★＝誤報
     ⇒★『書いてある』と『そこへ 繋いでいる』は 別★（前からの 決まり） */
  let 動く = text;
  try { 動く = 注記を外す(String(text)); } catch (e) { /* 外せなければ そのまま */ }
  const 動く所に在る = String(動く).indexOf(ref) >= 0;
  if (!動く所に在る) return { mark: '🟡', 訳: '注記（コメント）に 書いてあるだけ（繋いでいない）' };
  /* ★★試験の 中は 先に 見る★★（2026-09-07 実測で 踏んだ）
     ⇒ `nomiya-app-test/tests/supa-from-config.mjs` は
       ★名前に config が 入っている★ので「設定の ファイル」に 数えられた
     ⇒ 中身は `PROD_WAREHOUSE = "…"` ＝★本番では 走らせない為の 見張りの 定数★
     ⇒★★仲間を 敵と 数えた★★（指示役が 前に 名指しした 形と 同じ）
     ⇒★設定の ファイルは tests/ の 中には 置かない★ので 先に 外す */
  const 試験の中 = /(^|\/)(tests?|__tests__|spec)\//.test(String(path));
  if (!試験の中 && 設定のファイルか(path)) {
    return { mark: '🔴', 訳: '設定の ファイルの 中の 住所（画面が 読む 物）' };
  }
  if (繋ぐ呼び出しの中か(動く, ref)) return { mark: '🔴', 訳: '繋ぐ 呼び出しの 引数（fetch/createClient ほか）' };
  if (試験の期待値か(path, 動く, ref)) return { mark: '🟡', 訳: '試験の 期待値（not.toContain・名前→住所 の 対応表）' };
  return { mark: '🟡', 訳: '名前を 持っているだけ（役目が 読み取れない・要確認）' };
}

/** 見つかったrefと「向くべきref」から、〇×を決める。★見つからない=🟡（緑にしない）★ */
/** ④の判定（★純関数★＝倉庫を触らずに 自己確認できる）
 *  自分に … 自分が向いている倉庫の許可リストに 自分のURLが在るか（true/false/null=読めない）
 *  よそに … 向いていない倉庫の許可リストに 自分のURLが居るか（true=混ざり）
 *  ★2026-08-24 の決まり★：自分に在る＝当たり前／よそに居る＝★混ざり（赤）★ */
export function judgeRedirect(自分に, よそに) {
  if (自分に === null || よそに === null) return { mark: '🟡', text: '未測定（許可リストを読めない）' };
  if (自分に === false) return { mark: '🔴', text: '★自分の倉庫に無い＝忘れた人が戻って来られない' };
  if (よそに === true) return { mark: '🔴', text: '★混ざり＝別の環境の倉庫からも 戻れてしまう' };
  return { mark: '🟢', text: '自分の倉庫だけに在る（混ざり無し）' };
}

export function judge(found, want) {
  if (!found || found.length === 0) return { mark: '🟡', text: '未測定（向き先を読み取れない）' };
  const wrong = found.filter((r) => r !== want);
  if (wrong.length) return { mark: '🔴', text: '★別の倉庫を見ている: ' + wrong.join(',') };
  return { mark: '🟢', text: want === PROD_REF ? '本番' : 'テスト' };
}

/** 鍵の登録日が「引っ越しの日」より前か（前なら疑い） */
export function isStale(updatedAt, day = MIGRATION_DAY) {
  if (!updatedAt) return null;              // 分からない＝🟡
  return String(updatedAt).slice(0, 10) < day;
}

/* ══════════ ここから外に出る（全部 GET） ══════════ */
const argv = process.argv.slice(2);
const JSON_OUT = argv.includes('--json');

async function get(url, headers = {}, timeout = 25000) {
  try {
    const r = await fetch(url, { headers, redirect: 'follow', signal: AbortSignal.timeout(timeout) });
    return { ok: r.ok, status: r.status, text: await r.text() };
  } catch (e) {
    return { ok: false, status: 0, text: '', err: String(e && e.message).slice(0, 80) };
  }
}

/** 許可リストは鍵なしで測れる: 許可されていない戻り先は SITE_URL へ流される */
async function redirectAllowed(ref, target) {
  const u = `https://${ref}.supabase.co/auth/v1/verify?token=x&type=recovery&redirect_to=${encodeURIComponent(target)}`;
  try {
    const r = await fetch(u, { redirect: 'manual', signal: AbortSignal.timeout(25000) });
    const loc = r.headers.get('location') || '';
    if (!loc) return null;                                   // 測れなかった＝🟡
    return loc.split(/[?#]/)[0] === target;
  } catch (_) { return null; }
}

const rows = { c1: [], c2: [], c3: [], c4: [], c5: [], c6: [] };
let yellow = 0, red = 0;
/* ★★読めなかった ファイルの 本数★★（2026-09-07 指示役の 注文）
   ⇒ 印（🟢🔴🟡）の 数だけでは ★何本 見ていないか★が 表に 出ない
   ⇒ 上限に 当たって こぼれた 分は ★広げた つもりで 狭く なっている★ 所 */
const 読めず = { 上限で: 0, 木が切れた: 0, 取れなかった: 0 };
const bump = (mark) => { if (mark === '🟡') yellow++; if (mark === '🔴') red++; };

/* ── ① アプリのコード（＋アマかせは②の実効値も兼ねる） ── */
async function measureApps() {
  for (const a of APPS) {
    if (!a.host) { rows.c1.push({ ...a, mark: '—', text: a.note }); continue; }
    const r = await get(a.host + a.cfg);
    const found = r.ok ? refsIn(r.text) : [];
    const j = r.ok ? judge(found, a.want) : { mark: '🟡', text: `未測定（HTTP ${r.status}${r.err ? ' ' + r.err : ''}）` };
    bump(j.mark);
    rows.c1.push({ ...a, mark: j.mark, text: j.text, url: a.host + a.cfg });
  }
}

/* ── ② Vercel の環境変数 ── */
async function measureVercelEnv() {
  const tok = process.env.VERCEL_TOKEN;
  // 鍵が無くても「実効値」は分かる: /api/config を出しているアプリは配信で測れている
  for (const a of APPS.filter((x) => x.viaEnv)) {
    const c1 = rows.c1.find((r) => r.app === a.app && r.env === a.env);
    rows.c2.push({ app: a.app, env: a.env, mark: c1 ? c1.mark : '🟡',
      text: c1 ? `配信の /api/config で実効値を測った → ${c1.text}` : '未測定' });
  }
  if (!tok) {
    rows.c2.push({ app: '（その他の全プロジェクト）', env: '—', mark: '🟡',
      text: '未測定（VERCEL_TOKEN が無い）。2026-08-07 の手測りでは★アマかせ以外に倉庫のenvは無かった★が、増えていないかは測れていない' });
    yellow++;
    return;
  }
  const r = await get('https://api.vercel.com/v9/projects?limit=100&teamId=' + (process.env.VERCEL_TEAM_ID || ''),
    { Authorization: 'Bearer ' + tok });
  if (!r.ok) { rows.c2.push({ app: '（全プロジェクト）', env: '—', mark: '🟡', text: `未測定（Vercel API ${r.status}）` }); yellow++; return; }
  const projects = JSON.parse(r.text).projects || [];
  for (const p of projects) {
    const e = await get(`https://api.vercel.com/v9/projects/${p.id}/env?teamId=${process.env.VERCEL_TEAM_ID || ''}`,
      { Authorization: 'Bearer ' + tok });
    if (!e.ok) { rows.c2.push({ app: p.name, env: '—', mark: '🟡', text: `未測定（env ${e.status}）` }); yellow++; continue; }
    const envs = (JSON.parse(e.text).envs || []).filter((v) => /SUPABASE/i.test(v.key));
    if (!envs.length) { rows.c2.push({ app: p.name, env: '—', mark: '🟢', text: '倉庫のenv 無し（向き先はrepoのファイル）' }); continue; }
    // ★値は Encrypted で読めない＝「新しいか」しか言えない。だから🟡で出す★
    const old = envs.filter((v) => isStale(new Date(v.updatedAt || v.createdAt).toISOString()));
    rows.c2.push({ app: p.name, env: '—', mark: old.length ? '🟡' : '🟡',
      text: `倉庫のenv ${envs.length}本（値は読めない＝未測定）` +
        (old.length ? ` ★うち${old.length}本が ${MIGRATION_DAY} より前: ` + old.map((v) => `${v.key}/${v.target}`).join(' ') + '★' : '') });
    yellow++;
  }
}

/* ── ③ GitHub Actions の金庫 ── */
async function measureSecrets() {
  const gh = process.env.GH_SECRETS_TOKEN || process.env.GITHUB_TOKEN || process.env.GH_TOKEN;
  for (const b of BATCH_REPOS) {
    // まず「そのrepoのバッチが倉庫に触るのか」を workflow の中身で見る（公開repoなら鍵なしで読める）
    const wf = await ghGet(`https://api.github.com/repos/${b.repo}/contents/.github/workflows`, gh);
    let touches = null, files = [];
    if (wf.ok) {
      // ★この見張り自身(warehouse.yml)は数えない★
      //   読むだけの見張りが SUPABASE_ACCESS_TOKEN を受け取るので、
      //   除かないと「倉庫に触るバッチが有る」と★自分を指さして★しまう（実際に1回そうなった）。
      files = JSON.parse(wf.text).filter((f) => /\.ya?ml$/.test(f.name) && f.name !== 'warehouse.yml');
      touches = false;
      for (const f of files) {
        const c = await ghGet(f.download_url, gh);
        if (c.ok && /secrets\.SUPABASE/i.test(c.text)) { touches = true; break; }
      }
    }
    if (touches === false) { rows.c3.push({ repo: b.repo, mark: '🟢', text: `倉庫に触るバッチ 無し（workflow ${files.length}本を読んだ）` }); continue; }

    // 触るなら、鍵の登録日を見る（★値は読めない。日付でしか判断できない★）
    if (!process.env.GH_SECRETS_TOKEN) {
      rows.c3.push({ repo: b.repo, mark: '🟡',
        text: touches === null
          ? `未測定（workflowを読めない${b.private ? '・非公開repo' : ''}）`
          : '★倉庫に触るバッチが有る★ / 鍵の登録日は未測定（GH_SECRETS_TOKEN が無い）' });
      yellow++;
      continue;
    }
    const s = await ghGet(`https://api.github.com/repos/${b.repo}/actions/secrets`, process.env.GH_SECRETS_TOKEN);
    if (!s.ok) { rows.c3.push({ repo: b.repo, mark: '🟡', text: `未測定（secrets API ${s.status}）` }); yellow++; continue; }
    const secrets = (JSON.parse(s.text).secrets || []).filter((x) => /SUPABASE/i.test(x.name));
    if (!secrets.length) {
      rows.c3.push({ repo: b.repo, mark: '🟡', text: '★倉庫に触るバッチが有るのに 鍵が0本★＝バッチは何もできない（意図的ならその旨を書く）' });
      yellow++; continue;
    }
    const stale = secrets.filter((x) => isStale(x.updated_at));
    rows.c3.push({ repo: b.repo, mark: stale.length ? '🔴' : '🟡',
      text: `鍵 ${secrets.length}本` + (stale.length
        ? ` ★${MIGRATION_DAY} より前が ${stale.length}本＝引っ越し前の疑い: ` + stale.map((x) => `${x.name}(${x.updated_at.slice(0, 10)})`).join(' ') + '★'
        : `（全部 ${MIGRATION_DAY} 以降）／値は読めない＝実際の向き先は未測定`) });
    if (stale.length) red++; else yellow++;
  }
}

async function ghGet(url, token, timeout = 25000) {
  const h = { Accept: 'application/vnd.github+json', 'User-Agent': 'exally-warehouse-check' };
  if (token) h.Authorization = 'Bearer ' + token;
  return get(url, h, timeout);
}

/* ── ④ 認証の戻り先（★鍵が要らない★） ── */
async function measureRedirects() {
  /* ★向き先は 実物のコードから読む（表に手で書かない）★ */
  const 向き先 = new Map();
  for (const h of REDIRECT_HOSTS) {
    if (向き先.has(h.cfg)) continue;
    const r = await get(h.cfg);
    const found = r.ok ? refsIn(r.text) : [];
    向き先.set(h.cfg, found.includes(PROD_REF) ? PROD_REF : found.includes(TEST_REF) ? TEST_REF : null);
  }
  for (const h of REDIRECT_HOSTS) {
    const 自分 = 向き先.get(h.cfg);
    if (!自分) {
      bump('🟡');
      rows.c4.push({ ref: '—', app: h.app, url: h.url, mark: '🟡',
        text: '未測定（向き先を読めない＝どちらの倉庫の物か 決めない）' });
      continue;
    }
    const よそ = 自分 === PROD_REF ? TEST_REF : PROD_REF;
    const 自分に = await redirectAllowed(自分, h.url);
    const よそに = await redirectAllowed(よそ, h.url);
    const j = judgeRedirect(自分に, よそに);
    bump(j.mark);
    rows.c4.push({ ref: 自分 === PROD_REF ? '本番' : 'テスト', app: h.app, url: h.url, mark: j.mark, text: j.text });
  }
  // ★見張りが空振りしていないこと★: 許していない住所は必ず弾かれる
  for (const ref of [PROD_REF, TEST_REF]) {
    const ok = await redirectAllowed(ref, 'https://example.com/not-allowed.html');
    const mark = ok === null ? '🟡' : ok === false ? '🟢' : '🔴';
    bump(mark);
    rows.c4.push({ ref: ref === PROD_REF ? '本番' : 'テスト', app: '★空振り確認', url: 'https://example.com/…', mark,
      text: ok === false ? '許していない住所は弾かれた（測れている）' : ok ? '★何でも通っている＝許可リストが素通し★' : '未測定' });
  }
}

/* ── ⑤ Edge Function ── */
async function measureEdge() {
  const tok = process.env.SUPABASE_ACCESS_TOKEN;
  for (const ref of [PROD_REF, TEST_REF]) {
    const label = ref === PROD_REF ? '本番' : 'テスト';
    if (tok) {
      const r = await get(`https://api.supabase.com/v1/projects/${ref}/functions`, { Authorization: 'Bearer ' + tok });
      if (r.ok) {
        const fns = JSON.parse(r.text);
        const names = fns.map((f) => f.slug);
        const missing = EDGE_FUNCS.filter((n) => !names.includes(n));
        const mark = missing.length ? '🔴' : '🟢';
        bump(mark);
        rows.c5.push({ ref: label, mark,
          text: `${fns.length}本` + (missing.length ? ' ★足りない: ' + missing.join(',') : '（向き先は Deno.env の SUPABASE_URL＝自分の倉庫しか触れない）')
            + ' / 版: ' + fns.map((f) => `${f.slug}=v${f.version}`).join(' ') });
        continue;
      }
    }
    // 鍵が無い時の簡易判定: 関数URLを叩く。
    // ★2026-08-07 実測でここを直した★
    //   最初は「401なら居る」にしていたが、本番の dk-sync-jobs は GET に ★405★ を返す。
    //   鍵ありでは v12 で ACTIVE なのに ★居ない扱いで🔴になった（嘘の赤）★。
    //   ⇒ 「居ない」と言い切れるのは ★404だけ★。それ以外は「居る」。
    //   　 通信できなかった(0)は 🟡。版と中身は鍵が無ければ★どうやっても未測定★。
    const missing = [], unknown = [];
    for (const n of EDGE_FUNCS) {
      const r = await get(`https://${ref}.supabase.co/functions/v1/${n}`);
      if (r.status === 404) missing.push(n);
      else if (r.status === 0) unknown.push(n);
    }
    const mark = missing.length ? '🔴' : '🟡';
    bump(mark);
    rows.c5.push({ ref: label, mark,
      text: missing.length
        ? `★居ない: ${missing.join(',')}（関数URLが404）`
        : `${EDGE_FUNCS.length - unknown.length}/${EDGE_FUNCS.length}本が居る（404以外が返る）`
          + (unknown.length ? `／通信できず ${unknown.join(',')}` : '')
          + `／★版と中身は未測定（SUPABASE_ACCESS_TOKEN が無い）★` });
  }
}

/* ── ⑥ 見張りの道具そのもの ──
   ★「そのrepoが向くべき倉庫と違う物を持つ道具」を赤にする。理由のある物は★最初から★書いておく。
   　 後から慌てて足す形にすると、本物の直書きも一緒に通る（no-hardcoded-supa と同じ考え方）。
   mark:'🟢' = 設計として正しい ／ mark:'🟡' = 直すべきだが担当が他にある（★緑にはしない★） */
export const TOOL_ALLOWED = {
  'exally-zeroact/exally|tests/dbtest-seed.mjs': {
    mark: '🟢', why: 'DB-testに固定した手動ツール。本番refは「そこへ向いていたら即中止」の見張りとして持っている',
  },
  'exally-zeroact/exally-staging|tests/dbtest-seed.mjs': {
    mark: '🟢', why: '本番repoと同じ物。DB-testに固定した手動ツールで、本番refは「そこへ向いていたら即中止」の見張り',
  },
  'exally-zeroact/exally|scripts/check-warehouse-pointers.mjs': {
    mark: '🟢', why: '★この見張り自身★。本番とテストの両方のrefを「正解」として持たないと、何とも突き合わせられない',
  },
  'exally-zeroact/exally-staging|scripts/check-warehouse-pointers.mjs': {
    mark: '🟢', why: '同上（この見張り自身。両repoに同じ物を置く）',
  },
  'exally-zeroact/exally-staging|tests/pages-hosting.test.mjs': {
    mark: '🟢', why: 'わざと本番refを混ぜた作り物を通して、Pages配信の見張りが赤くなることを確かめる検査',
  },
  'exally-zeroact/exally-staging|kyuyo/scripts/pull-statutory.mjs': {
    mark: '🟢', why: '法定データ(最低賃金・保険料率)は★本番の中央倉庫が正★。テスト側から読んでも本番を見るのが設計',
  },
  'exally-zeroact/exally-staging|kyuyo/scripts/verify-statutory.mjs': { mark: '🟢', why: '同上（法定データは本番中央が正）' },
  'exally-zeroact/exally-staging|kyuyo/scripts/check-source-urls.mjs': { mark: '🟢', why: '同上（法定データは本番中央が正）' },
  'exally-zeroact/Daikou-app-test|scripts/check-hosts.mjs': {
    mark: '🟡', why: '★ダイコメの物。2026-08-07に指示役へ報告し、ダイコメセッションへ引き継ぎ済み。'
      + 'Exally側では直さない。直ったらこの行を消す★',
  },
  'exally-zeroact/Daikou-app-test|scripts/auth-mail-otp.mjs': {
    mark: '🟡', why: '★ダイコメの物。両対応へ直す作業が進行中（2026-08-07 時点で未コミット）。入ったらこの行を消す★',
  },
};

const TOOL_DIRS = /(^|\/)(scripts|tests|tools)\//;
/* ★設定の ファイルは どこに 在っても 読む★（2026-09-07）
   ⇒ 前は scripts/tests/tools しか 読んでいなかったので
     ★`js/dk-config.js` は ここに 一度も 来なかった★＝一番 大事な 1本 */
const 設定の道 = /(^|\/)[a-z0-9_-]*config[a-z0-9_-]*\.(m?js|ts|json)$/i;

/** ★★見る 範囲を 先に 数えて 書く★★（2026-09-07 指示役の 決まり）
 *
 *  ★なぜ★ … この 見張りは ★倉庫の 向き先★を 見る 物なのに、
 *    ★向き先が 書いてある `js/dk-config.js` を 一度も 読んでいませんでした★
 *    （読んでいたのは scripts / tests / tools だけ）
 *  ⇒★★『31個 見た』は『★見るべき 物を 見た★』では ない★★
 *  ⇒★見張りを 作ったら「何を 範囲に したか」を 先に 出し、
 *    ★一番 守りたい 物が その 範囲に 入っているか★を ★名指しで★ 確かめる★
 *
 *  @returns {{範囲:string, 入る:string[], 入らない:{道:string,訳:string}[]}}
 */
export function 見る範囲を数える(apps = APPS) {
  const 入る = [], 入らない = [];
  for (const a of apps) {
    if (!a.cfg) continue;
    const 道 = String(a.cfg).replace(/^\//, '');
    if (設定の道.test(道) || /\.html?$/i.test(道) || (TOOL_DIRS.test(道) && /\.(m?js|ts)$/.test(道))) {
      if (入る.indexOf(道) < 0) 入る.push(道);
    } else if (!入らない.some((x) => x.道 === 道)) {
      入らない.push({
        道: 道,
        訳: '配信の 口（repo の ファイルでは ない）',
      });
    }
  }
  return {
    範囲: 'scripts/tests/tools の js・ts ＋ ★config の 名前の ファイル★ ＋ ★画面(.html)★ は どこでも',
    入る: 入る, 入らない: 入らない,
  };
}
/* ★上限を 上げた（2026-09-07）★＝`.html` を 範囲に 足したので
   400 のままだと ★見ていない＝未測定★が 増えて 穴が 広がる */
const MAX_FILES = 900;
async function measureTools() {
  /* ★★見る 範囲を 先に 出す★★（2026-09-07 指示役の 決まり） */
  const 範 = 見る範囲を数える();
  rows.c6.push({ repo: '★見る 範囲★', mark: 範.入らない.length ? '🟡' : '🟢',
    text: 範.範囲
      + '／★一番 守りたい 物★ … 入る ' + 範.入る.length + '本（' + 範.入る.join(' ') + '）'
      + (範.入らない.length
        ? '／★範囲の 外 ' + 範.入らない.length + '本: '
          + 範.入らない.map((x) => x.道 + '（' + x.訳 + '）').join(' ') + '★'
        : '') });
  if (範.入らない.length) yellow++;

  const gh = process.env.GITHUB_TOKEN || process.env.GH_TOKEN;
  const usedAllow = new Set();
  for (const b of BATCH_REPOS) {
    const head = await ghGet(`https://api.github.com/repos/${b.repo}`, gh);
    if (!head.ok) { rows.c6.push({ repo: b.repo, mark: '🟡', text: `未測定（repoを読めない ${head.status}${b.private ? '・非公開' : ''}）` }); yellow++; continue; }
    const def = JSON.parse(head.text).default_branch || 'main';
    let tree = await ghGet(`https://api.github.com/repos/${b.repo}/git/trees/${def}?recursive=1`, gh, 60000);
    if (!tree.ok) tree = await ghGet(`https://api.github.com/repos/${b.repo}/git/trees/${def}?recursive=1`, gh, 60000); // 1回だけやり直す
    if (!tree.ok) { rows.c6.push({ repo: b.repo, mark: '🟡', text: `未測定（treeを読めない ${tree.status}${tree.err ? ' ' + tree.err : ''}）` }); yellow++; continue; }
    const t = JSON.parse(tree.text);
    const all = (t.tree || []).filter((x) => x.type === 'blob'
      && ((TOOL_DIRS.test(x.path) && /\.(m?js|ts)$/.test(x.path))
        || 設定の道.test(x.path)
        /* ★画面（.html）も 読む★＝住所を 直に 書いて いる 画面が 在る */
        || /\.html?$/i.test(x.path)));
    const look = all.slice(0, MAX_FILES);
    const hard = [], soft = [], known = [];
    for (const f of look) {
      const c = await ghGet(`https://raw.githubusercontent.com/${b.repo}/${def}/${f.path}`, gh);
      if (!c.ok) { 読めず.取れなかった++; continue; }
      /* ★役目で 分ける★（2026-09-07 指示役）＝「在る／無い」では 決めない */
      const 違う = [...new Set([...refsIn(c.text), ...bareRefsIn(c.text)])].filter((r) => r !== b.want);
      if (!違う.length) continue;
      const key = `${b.repo}|${f.path}`;
      const allow = TOOL_ALLOWED[key];
      if (allow) { usedAllow.add(key); known.push(`${f.path}[${allow.mark}]`); continue; }
      const 判 = 違う.map((r) => ({ r, ...役目で分ける(f.path, c.text, r) }));
      const 赤 = 判.filter((x) => x.mark === '🔴');
      if (赤.length) hard.push(`${f.path}(${赤.map((x) => x.r).join(',')}／${赤[0].訳})`);
      else soft.push(`${f.path}（${判[0].訳}）`);
    }
    const capped = all.length > look.length ? `／★${all.length - look.length}本は見ていない=未測定★` : '';
    if (all.length > look.length) 読めず.上限で += all.length - look.length;
    const truncated = t.truncated ? '／★treeが途中で切れている=未測定★' : '';
    if (t.truncated) 読めず.木が切れた++;
    const openHandover = known.some((k) => k.includes('🟡'));
    const mark = hard.length ? '🔴' : (soft.length || capped || truncated || openHandover) ? '🟡' : '🟢';
    bump(mark);
    rows.c6.push({ repo: b.repo, mark,
      text: `道具 ${look.length}/${all.length}本を読んだ${capped}${truncated}` +
        (hard.length ? ` ★別の倉庫へ【繋ぐ形】で書いている: ` + hard.slice(0, 6).join(' ') + '★' : '') +
        (soft.length ? `／★別の倉庫の名前を持つ(見張りの定数かもしれない・要確認) ${soft.length}本: ` + soft.slice(0, 6).join(' ') : '') +
        (known.length ? `／理由つきで許している ${known.length}本: ${known.join(' ')}` : '') });
    if (b === BATCH_REPOS[BATCH_REPOS.length - 1]) {
      // ★許可リストが現実から離れていないか（もう出てこない行が残っていないか）★
      const dead = Object.keys(TOOL_ALLOWED).filter((k) => !usedAllow.has(k));
      if (dead.length) {
        rows.c6.push({ repo: '（許可リストの点検）', mark: '🟡',
          text: '★もう出てこないのに残っている行がある＝消すこと: ' + dead.join(' ') });
        yellow++;
      } else {
        rows.c6.push({ repo: '（許可リストの点検）', mark: '🟢', text: `${Object.keys(TOOL_ALLOWED).length}行すべてが実際に出てきた（現実と合っている）` });
      }
    }
  }
}

/* ══════════ self-test（外へ出ない） ══════════ */
if (argv.includes('--self-test')) {
  let p = 0, f = 0;
  const T = (n, fn) => { try { fn(); p++; console.log('  ✓ ' + n); } catch (e) { f++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
  const eq = (a, b, m) => { if (JSON.stringify(a) !== JSON.stringify(b)) throw new Error(m + ' 実際=' + JSON.stringify(a)); };
  console.log('\n[warehouse-pointers] ★判定そのものが空振りしていないか★');
  T('URLの形から向き先を読める', () => eq(refsIn('x https://' + PROD_REF + '.supabase' + '.co y'), [PROD_REF], 'URL'));
  T('★裸の文字列は「繋ぐ形」とは別に数える（🔴にしないため）', () => {
    eq(refsIn("const PROD_REF = '" + PROD_REF + "';"), [], '繋ぐ形ではない');
    eq(bareRefsIn("const PROD_REF = '" + PROD_REF + "';"), [PROD_REF], '裸');
  });
  T('鍵(JWT)の中からも読める', () => {
    const pay = Buffer.from(JSON.stringify({ iss: 'supabase', ref: TEST_REF })).toString('base64url');
    eq(refsIn('k=eyJhbGciOiJIUzI1NiJ9.' + pay + '.sig'), [TEST_REF], 'JWT');
  });
  T('★読み取れない時は 🟢 にしない（🟡にする）', () => eq(judge([], PROD_REF).mark, '🟡', '空'));
  T('★違う倉庫を見ていたら 🔴', () => eq(judge([TEST_REF], PROD_REF).mark, '🔴', '誤接続'));
  T('正しければ 🟢', () => eq(judge([PROD_REF], PROD_REF).mark, '🟢', '正常'));
  T('★④ 自分の倉庫から1本 外すと 赤', () => eq(judgeRedirect(false, false).mark, '🔴', '自分に無い'));

  /* ══ ★役目で 分ける★（2026-09-07 指示役）══
     ★わざと 両方 作って ★🔴 と 🟡 が どちらも 出る★のを 見る★
     （前は「在る／無い」で 分けていて、★設定の ファイルを 見逃し／試験の 期待値を 誤報★した） */
  T('★🔴 設定の ファイルの 中の 住所（fetch の 引数では ない）', () => {
    const 中 = "const SUPA_URL = 'https://" + TEST_REF + ".supabase" + ".co';" + '\n' + "export const MY_SIDE = 'test';";
    const r = 役目で分ける('js/dk-config.js', 中, TEST_REF);
    eq(r.mark, '🔴', '設定の ファイル');
    if (!/設定/.test(r.訳)) throw new Error('訳が 出ていない: ' + r.訳);
  });
  T('★🔴 繋ぐ 呼び出しの 引数', () => {
    const 中 = "const c = createClient('https://" + TEST_REF + ".supabase" + ".co', key);";
    const r = 役目で分ける('scripts/x.mjs', 中, TEST_REF);
    eq(r.mark, '🔴', '繋ぐ 呼び出し');
    if (!/引数/.test(r.訳)) throw new Error('訳が 出ていない: ' + r.訳);
  });
  T('★🟡 試験の 期待値（not.toContain）', () => {
    const 中 = "expect(cfg).not.toContain('" + TEST_REF + "');";
    const r = 役目で分ける('tests/unit/a.test.js', 中, TEST_REF);
    eq(r.mark, '🟡', '期待値');
    if (!/期待値/.test(r.訳)) throw new Error('訳が 出ていない: ' + r.訳);
  });
  T('★🟡 名前→住所 の 対応表（試験の 中）', () => {
    const 中 = "const 表 = {" + '\n' + "  'Daikou-app-test': 'https://" + TEST_REF + ".supabase" + ".co'," + '\n' + "};";
    eq(役目で分ける('tests/unit/b.test.js', 中, TEST_REF).mark, '🟡', '対応表');
  });
  T('★★試験の 中でも 繋いでいたら 🔴★★（仲間だからと 甘くしない）', () => {
    const 中 = "await fetch('https://" + TEST_REF + ".supabase" + ".co/rest/v1/x');";
    eq(役目で分ける('tests/unit/c.test.js', 中, TEST_REF).mark, '🔴', '試験でも 繋げば 赤');
  });
  T('★役目が 読み取れない 時は 🟢 に しない（🟡）', () => {
    eq(役目で分ける('scripts/memo.mjs', "// メモ: " + TEST_REF, TEST_REF).mark, '🟡', '未確認');
  });
  T('★★注記に 書いてあるだけなら 🔴 に しない★★（実測で 踏んだ 誤報）', () => {
    const 中 = '/* テスト用DB(' + TEST_REF + ')とは 別の 倉庫を 指す */' + '\n'
      + "window.SUPA = { url: 'https://" + PROD_REF + ".supabase" + ".co' };";
    const r = 役目で分ける('js/supa-config.js', 中, TEST_REF);
    eq(r.mark, '🟡', '注記だけ');
    if (!/注記/.test(r.訳)) throw new Error('訳が 出ていない: ' + r.訳);
    /* ★同じ ファイルの 本物の 住所は 🔴★（甘くしすぎていない事） */
    eq(役目で分ける('js/supa-config.js', 中, PROD_REF).mark, '🔴', '本物の 住所');
  });
  T('★★見る 範囲に「一番 守りたい 物」が 入っている★★（2026-09-07 の 一番 大きい 抜け）', () => {
    const 範 = 見る範囲を数える();
    if (範.入る.indexOf('js/dk-config.js') < 0) throw new Error('dk-config.js が 範囲の 外');
    if (範.入る.indexOf('js/supa-config.js') < 0) throw new Error('supa-config.js が 範囲の 外');
    /* ★範囲の 外の 物は 隠さず 訳を 付けて 数える★（0件に しない） */
    for (const x of 範.入らない) if (!x.訳) throw new Error('範囲の 外なのに 訳が 無い: ' + x.道);
  });
  T('★★試験の 中の「config という 名前」を 設定と 数えない★★（実測で 踏んだ 誤報）', () => {
    const 中 = 'export const PROD_WAREHOUSE = "' + PROD_REF + '";'
      + ' if (url.includes(PROD_WAREHOUSE)) process.exit(1);';
    const r = 役目で分ける('tests/supa-from-config.mjs', 中, PROD_REF);
    eq(r.mark, '🟡', '見張りの 定数');
    /* ★でも 試験の 外なら 同じ 名前で 🔴★（甘くしすぎていない事） */
    eq(役目で分ける('js/supa-from-config.mjs', 中, PROD_REF).mark, '🔴', '試験の 外');
  });
  T('★★画面（.html）の 中の 直書きも 🔴★★（2026-09-07 指示役の 注文）', () => {
    const 中 = "<script>const SUPA_URL='https://" + TEST_REF + ".supabase" + ".co';</" + "script>";
    eq(役目で分ける('daikou-seikyu.html', 中, TEST_REF).mark, '🔴', '画面の 直書き');
    /* ★試験の 中の .html は 先に tests/ で 外れる★ */
    eq(役目で分ける('tests/fixtures/a.html', "expect(x).not.toContain('" + TEST_REF + "');", TEST_REF).mark,
      '🟡', '試験の 中');
  });
  T('★設定の ファイルの 見分け（名前だけで 決める）', () => {
    if (!設定のファイルか('js/dk-config.js')) throw new Error('dk-config.js を 設定と 見ていない');
    if (!設定のファイルか('src/config.ts')) throw new Error('config.ts を 設定と 見ていない');
    if (設定のファイルか('tests/unit/a.test.js')) throw new Error('試験を 設定と 見ている');
  });
  T('★④ 別の倉庫に1本 居たら 赤（混ざり）', () => eq(judgeRedirect(true, true).mark, '🔴', '混ざり'));
  T('★④ 自分だけに在れば 緑', () => eq(judgeRedirect(true, false).mark, '🟢', '正常'));
  T('★④ 読めない時は 🟢 にしない（🟡）', () => {
    eq(judgeRedirect(null, false).mark, '🟡', '自分側が読めない');
    eq(judgeRedirect(true, null).mark, '🟡', 'よそ側が読めない');
  });
  T('★④ 表に 向き先を手で書いていない（実物から読む）', () => {
    for (const h of REDIRECT_HOSTS) {
      if (!h.cfg) throw new Error(h.app + ': 向き先を読む場所が無い');
      if (h.want) throw new Error(h.app + ': ★表に手で書いている★');
    }
  });
  T('★引っ越しの日より前の鍵は「古い」と判定する', () => eq(isStale('2026-08-05T10:00:00Z'), true, '古い'));
  T('引っ越しの日以降の鍵は古くない', () => eq(isStale('2026-08-07T03:32:00Z'), false, '新しい'));
  T('★日付が分からない時は true/false を返さない（=🟡へ倒す）', () => eq(isStale(null), null, '不明'));
  T('★道具の許可リストは、全部に理由が書いてある', () => {
    for (const [k, v] of Object.entries(TOOL_ALLOWED)) {
      if (!v.why || v.why.length < 15) throw new Error(k + ': 理由が短すぎる');
      if (!['🟢', '🟡'].includes(v.mark)) throw new Error(k + ': mark が 🟢/🟡 以外');
      if (!k.includes('|')) throw new Error(k + ': repo|path の形になっていない');
    }
  });
  T('★引き継ぎ中の物は 🟡（緑にしていない）', () => {
    const h = TOOL_ALLOWED['exally-zeroact/Daikou-app-test|scripts/check-hosts.mjs'];
    eq(h && h.mark, '🟡', '引き継ぎ中');
  });
  console.log('\n' + p + ' passed, ' + f + ' failed');
  process.exit(f ? 1 : 0);
}

/* ══════════ 実行 ══════════
   ★import された時は 走らせない★（2026-08-29）
     PROD_REF / TEST_REF を 借りたい道具が 出てきた。見分けを 付けないと
     ★import しただけで 外へ GET を 何十本も 出し、最後の process.exit で
      借りた側が その場で 死ぬ★（実測＝apply-supabase-sql.mjs が 何も言わずに 終わっていた）。
     ＝ tests/run.js で 踏んだ物と 同じ型（あちらは require.main）。 */
const 走らせ役 = /check-warehouse-pointers\.mjs$/.test(String(process.argv[1] || '').split('\\').join('/'));
if (走らせ役) {
console.log('\n[warehouse-pointers] 倉庫の向き先 6か所 × 全アプリ（★読むだけ★）');
console.log(`  本番 = ${PROD_REF} / テスト = ${TEST_REF} / 引っ越しの日 = ${MIGRATION_DAY}\n`);

await measureApps();
await measureVercelEnv();
await measureSecrets();
await measureRedirects();
await measureEdge();
await measureTools();

const show = (title, list, fmt) => {
  console.log('\n══ ' + title + ' ══');
  list.forEach((r) => console.log('  ' + r.mark + ' ' + fmt(r)));
};
show('① アプリのコード（配信を叩いて実測）', rows.c1,
  (r) => `${(r.app + ' ' + r.env).padEnd(24)} ${r.text}${r.owner ? '  [担当:' + r.owner + ']' : ''}`);
show('② Vercel の環境変数', rows.c2, (r) => `${(r.app + ' ' + r.env).padEnd(34)} ${r.text}`);
show('③ GitHub Actions の金庫（バッチ）', rows.c3, (r) => `${r.repo.padEnd(42)} ${r.text}`);
show('④ 認証の戻り先 許可リスト（鍵不要で測れる）', rows.c4,
  (r) => `${r.ref.padEnd(6)} ${r.app.padEnd(18)} ${r.text}`);
show('⑤ Edge Function', rows.c5, (r) => `${r.ref.padEnd(6)} ${r.text}`);
show('⑥ 見張りの道具そのもの', rows.c6, (r) => `${r.repo.padEnd(42)} ${r.text}`);

const all = [...rows.c1, ...rows.c2, ...rows.c3, ...rows.c4, ...rows.c5, ...rows.c6];
console.log('\n── 実測 ──');
console.log(`  🟢 正しい : ${all.filter((r) => r.mark === '🟢').length}`);
console.log(`  🔴 ★誤り : ${all.filter((r) => r.mark === '🔴').length}★`);
console.log(`  🟡 ★未測定: ${all.filter((r) => r.mark === '🟡').length}★  ← 0件・異常なしにしない`);
console.log(`  —  対象外 : ${all.filter((r) => r.mark === '—').length}`);
/* ★★読めなかった 本数を 必ず 出す★★（2026-09-07 指示役）
   ⇒★『31個 見た』は『見るべき 物を 見た』では ない★
   ⇒★上限に 当たって こぼれた 分は「広げた つもりで 狭い」所★ */
{
  const 外 = 見る範囲を数える().入らない;
  const 合 = 読めず.上限で + 読めず.取れなかった;
  console.log(`  ★読めなかった ファイル : ${合}本★`
    + (合 ? `（上限で ${読めず.上限で} ／ 取れなかった ${読めず.取れなかった}）` : '')
    + (読めず.木が切れた ? `／★一覧が 途中で 切れた repo ${読めず.木が切れた}本★` : ''));
  console.log(`  ★見る 範囲の 外 : ${外.length}本★`
    + (外.length ? `（${外.map((x) => x.道 + '＝' + x.訳).join(' ／ ')}）` : ''));
}
if (JSON_OUT) console.log('\n' + JSON.stringify(rows, null, 2));
console.log(red ? '\n★🔴があります。向き先が違う所を直すこと★' : '\n🔴は0件（🟡の本数は上を見ること）');
process.exit(red ? 3 : 0);
}
