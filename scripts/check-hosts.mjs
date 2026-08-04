/* check-hosts.mjs — ★入口が生きているか／古い入口がちゃんと今の入口へ飛ぶか★
 *
 * なぜ必要か（2026-08-04）:
 *   司さんが古い入口(payslip-app-test)を「テスト用」だと思って開いた。
 *   中身は2026-08-02で止まっており、その日の直しが1つも入っていなかった。
 *   ★「古い」と書くだけでは人は開く。★ 塞いだ後も、塞がったままかを機械で見る。
 *
 * 見るもの（docs/HOSTS.md の表と1対1）:
 *   ① 今の入口が 200 で返るか
 *   ② 古い入口が、今の入口へ飛ぶか
 *      ・サーバ側の転送(301/302/308) … Location を見る
 *      ・ページ内の転送(GitHub Pages はサーバ転送が打てない) … 本文に飛び先が入っているかを見る
 *   ★どちらの形でも「飛び先が今の入口である」ことを確かめる（形は問わない）
 *
 * どこで回すか: ★通常CIには入れない★。外の都合(GitHub/Vercelのメンテ)で赤くなると、
 *   自分のせいでない赤で push が止まり、赤が信用されなくなる（source-urls と同じ理由）。
 *   → .github/workflows/hosts.yml で 週1(月曜9時JST)＋手動。
 *
 * 使い方: node scripts/check-hosts.mjs        （NGがあれば exit 3）
 *         node scripts/check-hosts.mjs --json
 */
const LIVE = [
  { name: '本番 ハブ', url: 'https://exally.vercel.app/hub.html' },
  { name: '本番 給与', url: 'https://exally.vercel.app/kyuyo/' },
  { name: '本番 グリッド', url: 'https://exally.vercel.app/book.html' },
  { name: 'テスト ハブ', url: 'https://exally-zeroact.github.io/exally-staging/hub.html' },
  { name: 'テスト 給与', url: 'https://exally-zeroact.github.io/exally-staging/kyuyo/' },
];

/* 古い入口 → どこへ飛ぶべきか。
   pending:true は「まだ塞いでいない（理由つき）」＝NGにはするが、何が残っているかが一覧で見える。 */
const OLD = [
  {
    name: '旧テスト 給与', url: 'https://exally-zeroact.github.io/payslip-app-test/',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
  },
  {
    name: '旧テスト 給与(直リンク)', url: 'https://exally-zeroact.github.io/payslip-app-test/meisai.html?t=X#h',
    to: 'https://exally-zeroact.github.io/exally-staging/kyuyo/',
    // ★飛び先の「うしろ」はブラウザが組み立てる（location.search/hash を足す）ので、
    //   HTMLの中には出てこない。だから「うしろを落とさない作りになっているか」を本文で見る。
    mustContain: ['location.search', 'location.hash'],
    why: '★うしろ(?t=/#)を落とさない作りか。落とすとWeb明細のリンクが死ぬ。'
      + '（実際に飛ぶことは実ブラウザで確認: /payslip-app-test/meisai.html?t=X#h → /exally-staging/kyuyo/meisai.html?t=X#h）',
  },
  {
    name: '旧本番 給与', url: 'https://payslip-app-olive.vercel.app/',
    to: 'https://exally.vercel.app/kyuyo/',
    pending: '★本番へ今回の直しが入ってから塞ぐ（先に塞ぐと飛び先が古いままになる）。2026-08-04時点は未対応。',
  },
];

const JSON_OUT = process.argv.includes('--json');

async function get(url) {
  try {
    const r = await fetch(url, { redirect: 'manual', headers: { 'User-Agent': 'Kyually-host-check/1.0' } });
    const loc = r.headers.get('location');
    let body = '';
    if (!loc && r.status < 400) body = await r.text().catch(() => '');
    else if (!loc) body = await r.text().catch(() => '');   // Pagesの404.html も中身を見る
    return { status: r.status, location: loc, body: body.slice(0, 200000) };
  } catch (e) {
    return { status: 0, error: e.message, body: '' };
  }
}

const results = { live: [], old: [] };

for (const h of LIVE) {
  const r = await get(h.url);
  const ok = r.status >= 200 && r.status < 400;
  results.live.push({ ...h, status: r.status, ok, why: ok ? null : ('HTTP ' + r.status + (r.error ? ' / ' + r.error : '')) });
}

for (const h of OLD) {
  const r = await get(h.url);
  let how = null, ok = false;
  if (r.location && r.location.indexOf(h.to) === 0) { how = 'サーバ転送 ' + r.status; ok = true; }
  else if (r.body && r.body.indexOf(h.to) >= 0) {
    how = 'ページ内転送'; ok = true;
    // 追加の条件（うしろを落とさない作りか 等）があれば、それも満たすこと
    for (const need of (h.mustContain || [])) {
      if (r.body.indexOf(need) < 0) { ok = false; how = 'ページ内転送(条件不足: ' + need + ')'; }
    }
  }
  results.old.push({
    ...h, status: r.status, how, ok: h.pending ? false : ok,
    why: h.pending ? h.pending : (ok ? null : '飛び先(' + h.to + ')が見つからない'),
  });
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
  results.old.forEach(x => console.log('  ' + (x.ok ? '✓' : '✗') + ' ' + String(x.status).padEnd(4) + ' ' + x.name.padEnd(22) + ' ' + (x.how || '—') + '\n      ' + x.url + '\n      → ' + x.to + (x.why ? '\n      ' + x.why : '')));
  console.log('\n── 実測 ──');
  console.log('  今の入口 OK ' + (results.live.length - ngLive.length) + ' / NG ' + ngLive.length);
  console.log('  古い入口 OK ' + (results.old.length - ngOld.length) + ' / NG ' + ngOld.length);
}

if (ngLive.length || ngOld.length) process.exitCode = 3;
