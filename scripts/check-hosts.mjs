/* check-hosts.mjs — ★入口が生きているか／古い入口がちゃんと今の入口へ飛ぶか★
 *
 * なぜ必要か（2026-08-04）:
 *   司さんが古い入口を「テスト用」だと思って開いた。中身は古く、その日の直しが1つも入っていなかった。
 *   ★「古い」と書くだけでは人は開く。★ 塞いだ後も、塞がったままかを機械で見る。
 *   旧本番 payslip-app-olive は ★本番のデータを見ている★ ので、開かれると事故になる。
 *
 * 見るもの（docs/HOSTS.md の表と1対1）:
 *   ① 今の入口が 200 で返るか
 *   ② 古い入口が、今の入口へ飛ぶか
 *      ・サーバ側の転送(301/302/308) … Location を見る
 *      ・ページ内の転送(GitHub Pages はサーバ転送が打てない) … 本文に飛び先が入っているかを見る
 *   ③ ★うしろ(?t=... / ?c=...)を落としていないか★
 *      落とすと Web明細のリンク（従業員に配ったQR）が死ぬ。
 *   ④ ★飛んだ先が本当に開けるか★（転送だけ成功して 404 に着地する事故を止める）
 *      旧ホストは cleanUrls で /meisai.html → /meisai に化ける。新ホストに /meisai は無い。
 *      ここを見ないと「転送は出来ている／でも真っ白」になる。
 *
 * どこで回すか: ★通常CIには入れない★。外の都合(Vercelのメンテ等)で赤くなると、
 *   自分のせいでない赤で push が止まり、赤が信用されなくなる（source-urls と同じ理由）。
 *   → .github/workflows/hosts.yml で 週1(月曜9時JST)＋手動。
 *   ★落ちた時に誰が見るか＝docs/HOSTS.md の「落ちた時に誰が見るか」に1行で書いてある。
 *
 * 使い方: node scripts/check-hosts.mjs
 *   ★終わり値★ … ★壊れて いる時だけ exit 3★
 *     ★予定どおりの 宿題（pending）は ★赤に しません★★（2026-09-16 に 分けました）
 *       ★訳★ 宿題で 赤に して いると ★本物の NG が 入っても 色が 動きません★
 *         node scripts/check-hosts.mjs --json
 *         node scripts/check-hosts.mjs --self-test   ★判定そのものが空振りしていないかを確かめる
 */

const LIVE = [
  { name: '本番 ハブ', url: 'https://exally.vercel.app/hub.html' },
  { name: '本番 グリッド', url: 'https://exally.vercel.app/book.html' },
  { name: 'テスト ハブ', url: 'https://exally-zeroact.github.io/exally-staging/hub.html' },
  // ★2026-08-07 いったん足して、同じ日に外した★
  //   テストの配信が2つ（github.io と Vercel）になっていたので見張りに載せたが、
  //   Vercel版は★git連携が無く、手で打った時だけ更新される★＝黙って古くなる形だった。
  //   司さんOKで★Vercel版のプロジェクトごと畳んだ★ので、見張る対象そのものが無くなった。
  //   ⇒ テストの配信は github.io の1本だけ（上の2行）。
];

/* 古い入口 → どこへ飛ぶべきか。
   mustKeep : 飛び先に必ず残っていないといけない文字（うしろを落としていないか）
   landing  : 飛んだ先を実際に叩いて 200 かを見る（転送だけ成功して404、を止める）
   pending  : まだ塞いでいない（理由つき）＝NGにするが、何が残っているかが一覧で見える */
const OLD = [
  /* ★★下の 5件（payslip-app-olive）は ★畳みました★★（2026-09-02 司さん OK）
       ★なぜ 一覧に 残すか★ … ★消すと「見張って いた」事ごと 消えます★
       ★なぜ NGに しないか★ … ★畳んだのだから 503 が ★正しい 姿★★
         ⇒★これを NG に して いると ★本物の NG が 入っても 色が 動きません★
         （2026-09-16 に 実測 … 古い入口 NG 7件 の うち ┅5件が これ★）
       ★決めた 訳★ 司さん「使わんもんは 消せや」（2026-09-02）
         ★repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★
       ★戻す 条件（back）★ … ★この ホストを もう 一度 立てる 日★
       ★見る 値打ちが 残って いる 所★
         ★万一 ★生き返った★ら 古い 明細が 開ける★⇒★200 が 返ったら 赤に します★ */
  {
    name: '旧本番 給与 トップ', url: 'https://payslip-app-olive.vercel.app/',
    folded: {
      why: '★畳んだ（2026-09-02 司さん「使わんもんは 消せや」）★＝★開かないのが 正しい 姿★。'
        + 'repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★。',
      back: '★この ホストを もう 一度 立てる 日★（その時は folded を 外して to を 入れる）',
    },
  },
  {
    name: '旧本番 Web明細(配布リンクの形)',
    folded: {
      why: '★畳んだ（2026-09-02 司さん「使わんもんは 消せや」）★＝★開かないのが 正しい 姿★。'
        + 'repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★。',
      back: '★この ホストを もう 一度 立てる 日★（その時は folded を 外して to を 入れる）',
    }, url: 'https://payslip-app-olive.vercel.app/meisai.html?t=PROBE&c=INIT',
    why: '★従業員に配ったQR/リンクの形。うしろ(?t=/?c=)を落とすと明細が開けなくなる。',
  },
  {
    name: '旧本番 Web明細(cleanUrlsで化けた形)',
    folded: {
      why: '★畳んだ（2026-09-02 司さん「使わんもんは 消せや」）★＝★開かないのが 正しい 姿★。'
        + 'repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★。',
      back: '★この ホストを もう 一度 立てる 日★（その時は folded を 外して to を 入れる）',
    }, url: 'https://payslip-app-olive.vercel.app/meisai?t=PROBE',
    why: '★旧ホストは cleanUrls で /meisai.html を /meisai に変えていた。新ホストに /meisai は無いので、'
      + '.html を付け直して飛ばさないと 404 に着地する。',
  },
  {
    name: '旧本番 管理',
    folded: {
      why: '★畳んだ（2026-09-02 司さん「使わんもんは 消せや」）★＝★開かないのが 正しい 姿★。'
        + 'repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★。',
      back: '★この ホストを もう 一度 立てる 日★（その時は folded を 外して to を 入れる）',
    }, url: 'https://payslip-app-olive.vercel.app/admin.html',
  },
  {
    name: '★旧本番 sw.js（ここだけ飛ばさない）',
    folded: {
      why: '★畳んだ（2026-09-02 司さん「使わんもんは 消せや」）★＝★開かないのが 正しい 姿★。'
        + 'repo は 凍結（給与の 本体は rakually）／★お客さんは 元から 1人も 見て いません★。',
      back: '★この ホストを もう 一度 立てる 日★（その時は folded を 外して to を 入れる）',
    }, url: 'https://payslip-app-olive.vercel.app/sw.js',
    expectNoRedirect: true, mustContainBody: ['unregister', 'caches.delete'],
    why: '★端末に住み着いた Service Worker は、サーバを塞いだだけでは消えない。'
      + 'sw.js まで転送すると SW の更新が失敗して★古いSWが永久に居座る★ので、ここだけは'
      + '「自分を登録解除してキャッシュを消す」中身をそのまま返す。',
  },
  // ★2026-08-07 「旧Exallyホーム(/)」と「(/home.html)」の2行を ここから外した★
  //   実測: /          → 307 で /daikou-seikyu.html（代行請求アプリ）へ
  //         /home.html → 404
  //   ★これは壊れたのではない。ダイコメ側の正しい変更★。
  //   `exally-test` は repo名変更(→ daikou-seikyu)で★ダイコメの製品になった★ので、
  //   このホストの入口をどうするかを決めるのも見るのも★ダイコメの担当★。
  //   Exally の見張りが他所の家の玄関を採点し続けると、
  //   ★自分では直せない赤が毎週鳴り、赤そのものが信用されなくなる★ので外す。
  //   （赤を消すために期待値を緩めたのではなく、★見る担当ごと手放した★という記録）
  {
    name: '★代行請求は塞がない（実務で動いている）', url: 'https://exally-test.vercel.app/daikou-seikyu.html',
    expectNoRedirect: true, mustContainBody: ['代行請求'],
    why: '★同じホストで実務が動いている。古いホームを塞ぐ時に巻き込んでいないことを、毎週ここで確かめる。'
      + '（2026-08-07: 持ち主はダイコメになったが、この1行は「Exally側の塞ぐ作業が巻き込んでいない」の確認なので残す）',
  },
  /* ★★下の 2件（GitHub Pages）は ★飛べません★★
       ★Pages は 3xx（サーバ転送）を 返せません★
       ⇒★「ここは 古い 場所です」の ★案内の 紙★を 返す 形が ★正しい 姿★★
       ★2026-09-16 に 中身を 測りました★（3,285バイト）
         ・★新しい 場所へ 案内して いる★
         ・★`location.search` と `location.hash` を 付け直して いる★
           ＝★従業員の QR・リンクの うしろ（?t= #）を 落として いません★
         ・★古い Service Worker を 登録解除し キャッシュを 消して いる★
       ★前は `to`（飛び先）が 無かった★
         ⇒★「飛び先(undefined)が 見つからない」で ずっと NG だった★
         ⇒★★見張りの 期待が 間違って いました（ホストは 壊れて いません）★★
       ★本当に 危ない 形★ … ★案内の 紙で は なく 古い 明細そのものが 開ける★
         ⇒★下の mustContainBody が それを 見ます★ */
  {
    name: '旧テスト 給与', url: 'https://exally-zeroact.github.io/payslip-app-test/',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
    mustContainBody: ['ここは古い場所です', 'location.search', 'location.hash'],
    michishirube: {
      why: '★わざと 残して ある 道しるべ★＝GitHub Pages は 3xx（サーバ転送）を 返せないので '
        + '「ここは 古い 場所です」の 紙を 返す。★消すと 従業員が 自分の 明細を 開けなく なります★。',
      back: '★古い QR・リンクを 持って いる 人が 居なく なったと 測れた 日★',
    },
  },
  {
    name: '旧テスト 給与(直リンク)', url: 'https://exally-zeroact.github.io/payslip-app-test/meisai.html?t=X#h',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
    michishirube: {
      why: '★わざと 残して ある 道しるべ★（上と 同じ 紙。Pages は 無い 道にも この 紙を 返す）',
      back: '★古い QR・リンクを 持って いる 人が 居なく なったと 測れた 日★',
    },
    // ★飛び先の「うしろ」はブラウザが組み立てる（location.search/hash を足す）ので、
    //   HTMLの中には出てこない。だから「うしろを落とさない作りになっているか」を本文で見る。
    mustContainBody: ['location.search', 'location.hash'],
  },
];

/* ★★2026-09-05 に 外しました★★（司さん「何のための Rakunally なんど」）
 *   ★給与(kyuyo/)を Exally から 消した★／★入口は 飛ばす（vercel.json の redirects）★
 *     /kyuyo/…  → https://rakually.vercel.app/kyuyo/…
 *     ★従業員が 持っている QR・リンクを 死なせない為★（?t= ?c= # は そのまま 付いて行く）
 *   ★AIが 読む 3本は lib/ へ 本体ごと 移した★（写しでは ない）
 *     shakaihoken-hyo / koyo-hoken / shouhizei-ritsu
 *   ★下の 一覧は 済んだ物★＝★消さずに 残す★（次に 同じ事を する 人の 為）
 */

/* ★これから塞ぐ入口（号令待ち）★ 2026-08-18 登録
 *   給与(kyuyo/)は Rakually へ移る。★22人が今 使っているので、移る日まで1バイトも外さない★。
 *   ここは「移す日に OLD へ移す物の一覧」＝★先に書いておく場所★。
 *   ★まだ塞いでいない物を赤にしない★。塞いでいないのは予定どおりなので、
 *   赤にすると自分のせいでない赤が毎週鳴り、★赤そのものが信用されなくなる★（pending と分ける理由）。
 *   ★数だけ毎回 出す★＝「一覧が在るのに誰も見ていない」を作らない。
 *   詳しい下調べ: docs/KYUYO_MOVE_INVENTORY.md
 */
const PLANNED = [
  {
    name: '★Web明細(配布リンクの形)', from: 'https://exally.vercel.app/kyuyo/meisai.html?t=…&c=…',
    to: '<Rakuallyの本番URL>/meisai.html',
    why: '★従業員に配ったQR/リンク。うしろ(?t= ?c= #)を落とすと明細が開けなくなる',
  },
  {
    name: '★Web明細(.html 無しの形)', from: 'https://exally.vercel.app/kyuyo/meisai',
    to: '<Rakuallyの本番URL>/meisai.html',
    why: '★.html を付け直さないと 404 に着地する（旧ホストで実際に踏んだ）',
  },
  {
    name: '★/sw.js は飛ばさない', from: 'https://exally.vercel.app/sw.js', to: '（飛ばさない）',
    why: '★端末に住み着いた Service Worker はサーバを塞いでも消えない。kyuyo/admin.html が /sw.js を登録している',
  },
];
/* ★先に行き先を決めないと外せない物★（2026-08-18 実測。docs/KYUYO_MOVE_INVENTORY.md の0章） */
const BLOCKERS = [
  '★済（2026-09-05）★ api/claude.js の3本 → lib/ へ本体ごと移した（shakaihoken-hyo / koyo-hoken / shouhizei-ritsu）',
  '★済（2026-09-05）★ hub.html のタイルを外した',
  '★★まだ★★ テスト線(exally-staging)の seikyu/ が kyuyo/lib/ を2本 読む（shiharai-chosho / shouhizei-ritsu）'
    + ' ⇒ ★本番から 消したので、テスト線を 同じにする 時に 先に 行き先を 決める★',
];

async function get(url) {
  try {
    const r = await fetch(url, { redirect: 'manual', headers: { 'User-Agent': 'Kyually-host-check/1.0' } });
    const loc = r.headers.get('location');
    const body = await r.text().catch(() => '');
    return { status: r.status, location: loc, body: body.slice(0, 200000) };
  } catch (e) {
    return { status: 0, error: e.message, body: '' };
  }
}
async function head(url) {
  try {
    const r = await fetch(url, { redirect: 'follow', headers: { 'User-Agent': 'Kyually-host-check/1.0' } });
    return r.status;
  } catch (e) { return 0; }
}

/* ★純関数: 1件の判定。self-test で作り物を通せる＝判定そのものが空振りしていないかを見る。 */
export function judge(h, r) {
  /* ★★畳んだ 入口★★（2026-09-16）
       ★開かないのが 正しい 姿★なので ★NG に しません★。
       ★一覧には 残します★（消すと「見張って いた」事ごと 消える）。
       ★ただし ★200 が 返ったら 赤★★＝★生き返って 古い 明細が 開けて いる★ */
  if (h.folded) {
    if (r.status === 200) {
      return { ok: false, how: 'HTTP 200', why: '★畳んだ はずの 入口が ★生き返って います★＝古い 明細が 開ける' };
    }
    return { ok: true, how: '畳んだ(HTTP ' + r.status + ')' };
  }
  if (h.expectNoRedirect) {
    if (r.location) return { ok: false, how: 'サーバ転送 ' + r.status, why: '★ここは飛ばしてはいけない（飛ばすとSWが更新できず居座る）' };
    if (r.status !== 200) return { ok: false, how: 'HTTP ' + r.status, why: '★中身が返っていない（SWのキルスイッチが消えている）' };
    for (const need of (h.mustContainBody || [])) {
      if (r.body.indexOf(need) < 0) return { ok: false, how: '中身あり', why: '★中身に「' + need + '」が無い＝キルスイッチになっていない' };
    }
    return { ok: true, how: '飛ばさない(200)' };
  }
  let how = null, ok = false;
  if (r.location && r.location.indexOf(h.to) === 0) { how = 'サーバ転送 ' + r.status; ok = true; }
  else if (r.body && r.body.indexOf(h.to) >= 0) { how = 'ページ内転送'; ok = true; }
  if (!ok) return { ok: false, how, why: '飛び先(' + h.to + ')が見つからない' };
  // ★うしろを落としていないか（サーバ転送は Location、ページ内転送は本文の作りを見る）
  const hay = r.location || r.body;
  for (const need of (h.mustKeep || [])) {
    if (hay.indexOf(need) < 0) return { ok: false, how, why: '★うしろを落としている（「' + need + '」が飛び先に無い）' };
  }
  for (const need of (h.mustContainBody || [])) {
    if (r.body.indexOf(need) < 0) return { ok: false, how, why: '★条件不足（本文に「' + need + '」が無い）' };
  }
  return { ok: true, how };
}

/* ══ self-test（判定そのものを、わざと壊して赤にする） ═══════════════════ */
if (process.argv.includes('--self-test')) {
  let pass = 0, fail = 0;
  const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };
  const H = { to: 'https://new/kyuyo/meisai.html', mustKeep: ['t=PROBE'] };
  console.log('\n[check-hosts --self-test] 判定そのものが空振りしていないか');
  T('★飛び先が違えば赤', () => { if (judge(H, { status: 308, location: 'https://old/other', body: '' }).ok) throw new Error('赤にならない'); });
  T('★うしろ(?t=)を落としたら赤', () => { const v = judge(H, { status: 308, location: 'https://new/kyuyo/meisai.html', body: '' }); if (v.ok) throw new Error('赤にならない'); });
  T('うしろが残っていれば緑', () => { const v = judge(H, { status: 308, location: 'https://new/kyuyo/meisai.html?t=PROBE', body: '' }); if (!v.ok) throw new Error('緑にならない: ' + v.why); });
  T('★転送そのものが無ければ赤（200のまま生きている）', () => { if (judge(H, { status: 200, location: null, body: '<html>給与</html>' }).ok) throw new Error('赤にならない'); });
  /* ★★畳んだ 入口★★（2026-09-16 に 足しました）
       ★気を 付ける 所★ … 「畳んだ」を ★何を しても 緑★に すると
         ★生き返って 古い 明細が 開けても 誰も 気づきません★
       ⇒★200 だけは 赤★ */
  const FOLD = { folded: '畳んだ（2026-09-02 司さん OK）' };
  T('畳んだ 入口は 503 で 緑（★開かないのが 正しい 姿★）', () => {
    if (!judge(FOLD, { status: 503, location: null, body: '' }).ok) throw new Error('緑に ならない');
  });
  T('畳んだ 入口は 404 でも 緑', () => {
    if (!judge(FOLD, { status: 404, location: null, body: '' }).ok) throw new Error('緑に ならない');
  });
  T('★畳んだ はずの 入口が 200 で 生き返ったら 赤', () => {
    if (judge(FOLD, { status: 200, location: null, body: '<html>明細</html>' }).ok) throw new Error('赤に ならない');
  });
  /* ★★案内の 紙（GitHub Pages は 3xx を 返せない）★★ */
  const PG = { to: 'https://new/kyuyo/', mustContainBody: ['ここは古い場所です', 'location.search'] };
  T('案内の 紙が 新しい 場所を 指し うしろも 残すなら 緑', () => {
    const v = judge(PG, { status: 200, location: null,
      body: 'ここは古い場所です https://new/kyuyo/ location.search' });
    if (!v.ok) throw new Error('緑に ならない: ' + v.why);
  });
  T('★案内の 紙が うしろ（?t= #）を 落とす 作りなら 赤', () => {
    const v = judge(PG, { status: 200, location: null,
      body: 'ここは古い場所です https://new/kyuyo/' });
    if (v.ok) throw new Error('赤に ならない');
  });
  T('★古い 明細そのものが 開けて いたら 赤（案内の 紙で ない）', () => {
    const v = judge(PG, { status: 200, location: null,
      body: '<html>給与明細 https://new/kyuyo/ location.search</html>' });
    if (v.ok) throw new Error('赤に ならない');
  });
  const SW = { expectNoRedirect: true, mustContainBody: ['unregister'] };
  T('★sw.js が飛ばされていたら赤', () => { if (judge(SW, { status: 308, location: 'https://new/', body: '' }).ok) throw new Error('赤にならない'); });
  T('★sw.js の中身がキルスイッチでなければ赤', () => { if (judge(SW, { status: 200, location: null, body: 'self.addEventListener("fetch",...)' }).ok) throw new Error('赤にならない'); });
  T('sw.js がキルスイッチのままなら緑', () => { if (!judge(SW, { status: 200, location: null, body: 'registration.unregister()' }).ok) throw new Error('緑にならない'); });
  /* ★これから塞ぐ入口の一覧が、黙って空になったり 二重管理になったりしないよう見張る★ */
  T('★これから塞ぐ入口の一覧が空になっていない', () => { if (!PLANNED.length) throw new Error('一覧が空＝誰も見ていない状態'); });
  T('★これから塞ぐ入口が「古い入口」にも入っていない（二重管理しない）', () => {
    const olds = new Set(OLD.map(o => o.url.split('?')[0]));
    const dup = PLANNED.filter(p => olds.has(p.from.split('?')[0]));
    if (dup.length) throw new Error('両方に居る: ' + dup.map(d => d.from).join(','));
  });
  T('★Web明細は「うしろを落とさない」と「.html を付け直す」の両方が一覧に在る', () => {
    const t = PLANNED.map(p => p.from + ' ' + (p.why || '')).join('|');
    if (!/\?t=/.test(t)) throw new Error('うしろ(?t=)の形が一覧に無い');
    if (!PLANNED.some(p => /\/meisai$/.test(p.from))) throw new Error('.html 無しの形が一覧に無い');
  });
  T('★先に行き先を決める物の一覧が空になっていない', () => { if (!BLOCKERS.length) throw new Error('一覧が空'); });
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exitCode = fail ? 1 : 0;
} else {
  /* ══ 本番（実物を叩く） ═══════════════════════════════════════════════ */
  const JSON_OUT = process.argv.includes('--json');
  const results = { live: [], old: [] };

  for (const h of LIVE) {
    const r = await get(h.url);
    const ok = r.status >= 200 && r.status < 400;
    results.live.push({ ...h, status: r.status, ok, why: ok ? null : ('HTTP ' + r.status + (r.error ? ' / ' + r.error : '')) });
  }

  for (const h of OLD) {
    const r = await get(h.url);
    let v = judge(h, r);
    let landStatus = null;
    if (v.ok && h.landing && r.location) {
      landStatus = await head(r.location);
      if (landStatus !== 200) v = { ok: false, how: v.how, why: '★飛んだ先が開けない（HTTP ' + landStatus + '）＝転送は出来ているのに真っ白になる' };
    }
    if (h.pending) v = { ok: false, how: v.how, why: h.pending };
    results.old.push({ ...h, status: r.status, location: r.location, landStatus, ...v });
  }

  const ngLive = results.live.filter(x => !x.ok);
  /* ★★「壊れて いる」と 「予定どおりの 宿題」を 分けます★★（2026-09-16）
       ★前は どちらも exit 3（赤）だった★
       ⇒★★宿題が 1件 残って いる 間は ずっと 赤★★
       ⇒★明日「今の 入口 NG 3」に なっても 色は 赤のまま＝★誰も 気づきません★
       ★同じ 型★ 本番の WebKit が 2日 赤のままだった（その間 本番へ 3回 押して いた）
       ⇒★★色は 1つの 事だけを 言う★★
         ・★お客さんに 出る 所が 壊れて いる★ ⇒ ★赤★
         ・★予定どおりの 宿題★         ⇒ ★緑 ＋ 残り件数を 出す★（★色では 言わない★） */
  /* ★★古い 入口を ┅3つの 印★に 分けます★★（2026-09-16）
       ㆁ★畳んだ（folded）★         … 開かないのが 正しい 姿（┅5件★）
       ㆂ★わざと 残す 道しるべ（michishirube）★ … ★消すと 明細が 開けなく なる★（┅2件★）
       ㆃ★本当の 宿題（pending）★  … ★まだ 塩いで いない★（今 ┅0件★）
     ★なぜ 1つの 山に しないか★
       2026-09-16 に ★古い入口 NG 7件★ と 出て いましたが
       ★中身を 測ったら 7件とも 「正しい 姿」でした★。
       ⇒★★本物の NG が 入っても 色が 動かない 形★★だった
     ★印は 全部 訳(why) と 戻す 条件(back) 付き★（`tests-no-ci.json` と 同じ 形） */
  const 宿題 = results.old.filter(x => !x.ok && x.pending);
  const ngOld = results.old.filter(x => !x.ok && !x.pending);

  if (JSON_OUT) {
    console.log(JSON.stringify({ ngLive: ngLive.length, ngOld: ngOld.length, 宿題: 宿題.length, results }, null, 1));
  } else {
    console.log('\n[check-hosts] 入口の生死と、古い入口の飛び先（docs/HOSTS.md と1対1）\n');
    console.log('■ 今の入口');
    results.live.forEach(x => console.log('  ' + (x.ok ? '✓' : '✗') + ' ' + String(x.status).padEnd(4) + ' ' + x.name.padEnd(14) + ' ' + x.url));
    console.log('\n■ 古い入口（今の入口へ飛ぶか）');
    results.old.forEach(x => console.log('  ' + (x.ok ? '✓' : '✗') + ' ' + String(x.status).padEnd(4) + ' ' + x.name
      + '\n      ' + x.url
      + (x.location ? '\n      → ' + x.location + (x.landStatus ? '  [飛び先 HTTP ' + x.landStatus + ']' : '') : (x.to ? '\n      → ' + x.to : ''))
      + '\n      ' + (x.how || '—') + (x.why ? '  ／ ' + x.why : '')));
    console.log('\n■ これから塞ぐ入口（★号令待ち。まだ塞いでいないので赤にしない★）');
    PLANNED.forEach(x => console.log('  ・' + x.name + '\n      ' + x.from + '\n      → ' + x.to
      + (x.why ? '\n      ' + x.why : '')));
    console.log('  ★先に行き先を決めないと外せない物 ' + BLOCKERS.length + '件★');
    BLOCKERS.forEach(b => console.log('      - ' + b));

    console.log('\n── 実測 ──');
    console.log('  今の入口 OK ' + (results.live.length - ngLive.length) + ' / NG ' + ngLive.length);
      const 畳んだ数 = results.old.filter(x => x.folded).length;
    const 道しるべ数 = results.old.filter(x => x.michishirube).length;
    console.log('  古い入口 OK ' + (results.old.length - ngOld.length - 宿題.length)
      + ' / ★NG ' + ngOld.length + '★'
      + ' / 予定どおりの 宿題 ' + 宿題.length + '件（★色には しません★）');
    宿題.forEach(x => console.log('      ・' + x.name + ' … ' + x.pending));
    /* ★★印の 内訳を 必ず 出します★★
         ★「OK 8」だけ 出すと ★何が なぜ OK なのか 分かりません★ */
    console.log('    内訳 … ★畳んだ ' + 畳んだ数 + '件★（開かないのが 正しい 姿）'
      + ' ／ ★わざと 残す 道しるべ ' + 道しるべ数 + '件★（消すと 明細が 開けなく なる）'
      + ' ／ 本当の 宿題 ' + 宿題.length + '件');
    console.log('  これから塞ぐ入口 ' + PLANNED.length + '件（★まだ0件も塞いでいない＝予定どおり★）'
      + ' / 先に決める物 ' + BLOCKERS.length + '件');
  }

  /* ★★終わり値は 「お客さんに 出る 所が 壊れて いるか」だけで 決めます★★
       ★今の 入口 NG★     … お客さんが 今 開けない
       ★古い 入口 NG★     … お客さんが 古い リンクで 飛べない（★これも 客に 出ます★）
       ★予定どおりの 宿題★ … ★まだ 塩いで いないだけ★＝★壊れて は いません★
       ⇒★宿題で 赤に すると ★壊れた 日に 色が 動きません★★ */
  if (ngLive.length || ngOld.length) process.exitCode = 3;
}
