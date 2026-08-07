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
 * 使い方: node scripts/check-hosts.mjs        （NGがあれば exit 3）
 *         node scripts/check-hosts.mjs --json
 *         node scripts/check-hosts.mjs --self-test   ★判定そのものが空振りしていないかを確かめる
 */

const LIVE = [
  { name: '本番 ハブ', url: 'https://exally.vercel.app/hub.html' },
  { name: '本番 給与', url: 'https://exally.vercel.app/kyuyo/' },
  { name: '本番 Web明細', url: 'https://exally.vercel.app/kyuyo/meisai.html?t=PROBE' },
  { name: '本番 グリッド', url: 'https://exally.vercel.app/book.html' },
  { name: 'テスト ハブ', url: 'https://exally-zeroact.github.io/exally-staging/hub.html' },
  { name: 'テスト 給与', url: 'https://exally-zeroact.github.io/exally-staging/kyuyo/' },
  // ★2026-08-07 追加★ テストの配信は2つある（github.io と Vercel）。
  //   Vercel版は★git連携が無く、手で打った時だけ更新される★＝黙って古いまま生き続ける。
  //   見張っていなかったので、ここに載せて毎週 目に入るようにする。
  //   （どちらに寄せるかは未決。決まるまでは両方 生きている前提で見る）
  { name: 'テスト ハブ(Vercel版・手打ち更新)', url: 'https://exally-staging.vercel.app/hub.html' },
];

/* 古い入口 → どこへ飛ぶべきか。
   mustKeep : 飛び先に必ず残っていないといけない文字（うしろを落としていないか）
   landing  : 飛んだ先を実際に叩いて 200 かを見る（転送だけ成功して404、を止める）
   pending  : まだ塞いでいない（理由つき）＝NGにするが、何が残っているかが一覧で見える */
const OLD = [
  {
    name: '旧本番 給与 トップ', url: 'https://payslip-app-olive.vercel.app/',
    to: 'https://exally.vercel.app/kyuyo/', landing: true,
  },
  {
    name: '旧本番 Web明細(配布リンクの形)', url: 'https://payslip-app-olive.vercel.app/meisai.html?t=PROBE&c=INIT',
    to: 'https://exally.vercel.app/kyuyo/meisai.html', mustKeep: ['t=PROBE', 'c=INIT'], landing: true,
    why: '★従業員に配ったQR/リンクの形。うしろ(?t=/?c=)を落とすと明細が開けなくなる。',
  },
  {
    name: '旧本番 Web明細(cleanUrlsで化けた形)', url: 'https://payslip-app-olive.vercel.app/meisai?t=PROBE',
    to: 'https://exally.vercel.app/kyuyo/meisai.html', mustKeep: ['t=PROBE'], landing: true,
    why: '★旧ホストは cleanUrls で /meisai.html を /meisai に変えていた。新ホストに /meisai は無いので、'
      + '.html を付け直して飛ばさないと 404 に着地する。',
  },
  {
    name: '旧本番 管理', url: 'https://payslip-app-olive.vercel.app/admin.html',
    to: 'https://exally.vercel.app/kyuyo/admin.html', landing: true,
  },
  {
    name: '★旧本番 sw.js（ここだけ飛ばさない）', url: 'https://payslip-app-olive.vercel.app/sw.js',
    expectNoRedirect: true, mustContainBody: ['unregister', 'caches.delete'],
    why: '★端末に住み着いた Service Worker は、サーバを塞いだだけでは消えない。'
      + 'sw.js まで転送すると SW の更新が失敗して★古いSWが永久に居座る★ので、ここだけは'
      + '「自分を登録解除してキャッシュを消す」中身をそのまま返す。',
  },
  // ★2026-08-07 この2本が赤に変わった（実測）★
  //     https://exally-test.vercel.app/          → 307 で /daikou-seikyu.html（自分のアプリ）へ
  //     https://exally-test.vercel.app/home.html → 404
  //   このホストは repo名変更(exally-test → daikou-seikyu)で★ダイコメの製品になった★ので、
  //   飛び先を決めるのは向こうの担当。★赤を消すために期待値を書き換えない★
  //   （測り方を実態に合わせて緩めると、次に本当に壊れた時に気づけない）。
  //   ダイコメ側で「Exallyのハブへ戻す／戻さない」が決まったら、その時に この2行を直す。
  {
    name: '旧Exallyホーム(/)', url: 'https://exally-test.vercel.app/',
    to: 'https://exally.vercel.app/hub.html', landing: true,
    why: '★入口が1つ増えているだけで価値がゼロ。司さんが古い入口を「テスト用」と誤解した前科がある。'
      + '★2026-08-07: 実際には代行請求アプリへ307。ホストの持ち主はダイコメ＝Exally側では直さない。',
  },
  {
    name: '旧Exallyホーム(/home.html)', url: 'https://exally-test.vercel.app/home.html',
    to: 'https://exally.vercel.app/hub.html', landing: true,
    why: '★2026-08-07: 404 になった。ホストの持ち主はダイコメ＝Exally側では直さない。',
  },
  {
    name: '★代行請求は塞がない（実務で動いている）', url: 'https://exally-test.vercel.app/daikou-seikyu.html',
    expectNoRedirect: true, mustContainBody: ['代行請求'],
    why: '★同じホストで実務が動いている。古いホームを塞ぐ時に巻き込んでいないことを、毎週ここで確かめる。',
  },
  {
    name: '旧テスト 給与', url: 'https://exally-zeroact.github.io/payslip-app-test/',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
  },
  {
    name: '旧テスト 給与(直リンク)', url: 'https://exally-zeroact.github.io/payslip-app-test/meisai.html?t=X#h',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
    // ★飛び先の「うしろ」はブラウザが組み立てる（location.search/hash を足す）ので、
    //   HTMLの中には出てこない。だから「うしろを落とさない作りになっているか」を本文で見る。
    mustContainBody: ['location.search', 'location.hash'],
  },
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
  const SW = { expectNoRedirect: true, mustContainBody: ['unregister'] };
  T('★sw.js が飛ばされていたら赤', () => { if (judge(SW, { status: 308, location: 'https://new/', body: '' }).ok) throw new Error('赤にならない'); });
  T('★sw.js の中身がキルスイッチでなければ赤', () => { if (judge(SW, { status: 200, location: null, body: 'self.addEventListener("fetch",...)' }).ok) throw new Error('赤にならない'); });
  T('sw.js がキルスイッチのままなら緑', () => { if (!judge(SW, { status: 200, location: null, body: 'registration.unregister()' }).ok) throw new Error('緑にならない'); });
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
  const ngOld = results.old.filter(x => !x.ok);

  if (JSON_OUT) {
    console.log(JSON.stringify({ ngLive: ngLive.length, ngOld: ngOld.length, results }, null, 1));
  } else {
    console.log('\n[check-hosts] 入口の生死と、古い入口の飛び先（docs/HOSTS.md と1対1）\n');
    console.log('■ 今の入口');
    results.live.forEach(x => console.log('  ' + (x.ok ? '✓' : '✗') + ' ' + String(x.status).padEnd(4) + ' ' + x.name.padEnd(14) + ' ' + x.url));
    console.log('\n■ 古い入口（今の入口へ飛ぶか）');
    results.old.forEach(x => console.log('  ' + (x.ok ? '✓' : '✗') + ' ' + String(x.status).padEnd(4) + ' ' + x.name
      + '\n      ' + x.url
      + (x.location ? '\n      → ' + x.location + (x.landStatus ? '  [飛び先 HTTP ' + x.landStatus + ']' : '') : (x.to ? '\n      → ' + x.to : ''))
      + '\n      ' + (x.how || '—') + (x.why ? '  ／ ' + x.why : '')));
    console.log('\n── 実測 ──');
    console.log('  今の入口 OK ' + (results.live.length - ngLive.length) + ' / NG ' + ngLive.length);
    console.log('  古い入口 OK ' + (results.old.length - ngOld.length) + ' / NG ' + ngOld.length);
  }

  if (ngLive.length || ngOld.length) process.exitCode = 3;
}
