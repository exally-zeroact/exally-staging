/* check-deployed-version.mjs — ★配信されている物が、今のコードと同じ版か★
 *
 * なぜ必要か（2026-08-18 に実際に起きた）:
 *   直しを push し、CI(GitHub Actions)も緑になった。★それでも配信は前の版のままだった。★
 *   GitHub → Vercel の合図(webhook)が届かず、★ビルドが失敗したのではなく 始まってすらいなかった★。
 *   その間、客の画面では 実物の式 15,126本のうち 12,245本(81%)が #ERROR のままだった。
 *   ★「push した」「CIが緑」「repoに入っている」は、どれも「客に届いた」ではない。★
 *   同じ型を2026-08-07 にも踏んでいる（git連携の無い配信が黙って古くなった）。
 *
 * 何を見るか
 *   ① 配信されている画面の ?v=<刻印> が、★今のコードの刻印★と同じか
 *      （刻印は全JS/CSSの内容から作る決定論的な値＝1文字でも違えば変わる）
 *   ② ★遅れて読む部品まで含めて、配信に実在するか★
 *      book.html は lib/table-refs.js などを ★後から★ 読む。
 *      HTMLだけ新しくても、部品が配信に無ければ ★押した時に初めて 404 で死ぬ★。
 *      実際 2026-08-18 は lib/table-refs.js が 404 だった。
 *
 * どこで回すか: ★通常CI(ci.yml)には入れない★。外の都合（Vercelの混雑・配信の途中）で赤くなると
 *   自分のせいでない赤で push が止まり、赤そのものが信用されなくなる（check-hosts と同じ理由）。
 *   → .github/workflows/hosts.yml で 週1＋手動。
 *   ★push の直後は「まだ配信されていない」で赤になるのが正しい★（それが見たい物）。
 *
 * 使い方: node scripts/check-deployed-version.mjs
 *         node scripts/check-deployed-version.mjs --host https://exally.vercel.app
 *         node scripts/check-deployed-version.mjs --json
 *         node scripts/check-deployed-version.mjs --self-test   ★判定が空振りしていないか（外へ出ない）
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { buildHash } from './stamp-build.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/* ★このリポジトリが養っている配信はどれか★
 *   js/env-badge.js は「テスト環境の帯」＝★テスト線にしか置かない物★（本番に出さないと決めてある）。
 *   これで本番かテスト線かが分かる＝同じファイルを両方の repo に置ける（片方だけに入れない）。 */
const IS_STAGING = fs.existsSync(path.join(ROOT, 'js', 'env-badge.js'));
const DEFAULT_HOST = IS_STAGING
  ? 'https://exally-zeroact.github.io/exally-staging'
  : 'https://exally.vercel.app';

const PAGES = ['book.html', 'hub.html'];
/* 部品の実在まで見る画面（★遅れて読む物が多い＝壊れても押すまで分からない★） */
const ASSET_PAGE = 'book.html';

/** 画面のHTMLから、読んでいる部品を全部 拾う（静的な src と、遅れて読む物の両方） */
export function assetsOf(html) {
  const out = new Set();
  for (const m of html.matchAll(/<script src="([^"]+)"><\/script>/g)) {
    const s = m[1].split('?')[0];
    if (!/^https?:/.test(s)) out.add(s);
  }
  for (const m of html.matchAll(/<link[^>]+href="([^"]+\.css)(?:\?[^"]*)?"/g)) {
    if (!/^https?:/.test(m[1])) out.add(m[1].split('?')[0]);
  }
  /* ★遅れて読む物＝_loadScript('lib/….js' + v) の形。ここを見ないと 404 を見逃す★ */
  for (const m of html.matchAll(/_loadScript\(\s*'([^']+\.(?:js|css))'/g)) out.add(m[1]);
  return [...out].sort();
}

/** 配信されたHTMLに貼られている刻印（?v=…）を全部 取る */
export function stampsOf(html) {
  const s = new Set();
  for (const m of html.matchAll(/\?v=([0-9a-f]{6,})/g)) s.add(m[1]);
  return [...s];
}

/** 判定（★純関数＝self-test でわざと壊せる★） */
export function judgeVersion(want, stamps) {
  if (!stamps.length) return { ok: false, why: '配信に刻印(?v=)が1つも無い＝古い作りか、別の物が返っている' };
  const bad = stamps.filter(v => v !== want);
  if (bad.length) {
    return { ok: false, why: '★配信が今のコードと違う版★ 配信=' + stamps.join(',') + ' / 今のコード=' + want };
  }
  return { ok: true, why: null };
}

/* ★HTTP 200 は「そのファイルが在る」ではない★（2026-08-18 に自分で踏んだ）
   保護のかかった配信は ★どのパスでも 302 → ログイン画面 200★ を返す。
   redirect を追って status だけ見ると、.js が1本も無くても ★12本ぜんぶ OK★ になった。
   ⇒ ① redirect は追わない ② 中身の署名（元のファイルの頭）が入っているかまで見る
      ③ 保護で読めない物は ★🟡未測定★ にする（緑にも赤にもしない） */
const SSO_RE = /_vercel\/sso|vercel\.com\/sso-api|sso\.vercel\.com|Redirecting/i;

export function judgeAsset(res, sig) {
  if (res.status === 0) return { state: 'ng', why: '繋がらない' + (res.error ? '（' + res.error + '）' : '') };
  if (res.status >= 300 && res.status < 400) {
    const loc = res.location || '';
    if (SSO_RE.test(loc)) return { state: 'unknown', why: '★保護がかかっていて中身を読めない（未測定）' };
    return { state: 'ng', why: 'HTTP ' + res.status + ' で飛ばされる → ' + loc };
  }
  if (res.status !== 200) return { state: 'ng', why: 'HTTP ' + res.status };
  if (SSO_RE.test(res.body.slice(0, 400))) return { state: 'unknown', why: '★保護の画面が返っている（未測定）' };
  const ct = String(res.contentType || '');
  if (/\.js$/.test(sig.name) && /text\/html/i.test(ct)) return { state: 'ng', why: '中身がHTML（' + ct + '）＝別の物が返っている' };
  /* ★比べる前に改行をそろえる★（CRLF/LF の差で本物の差が見かけの差に埋もれる／逆に嘘の赤が出る） */
  const nl = s => String(s).replace(/\r\n/g, '\n');
  if (sig.head && nl(res.body).indexOf(nl(sig.head)) < 0) {
    return { state: 'ng', why: '★中身がうちのファイルと違う★（頭の署名が入っていない）' };
  }
  return { state: 'ok', why: null };
}

async function get(url, noFollow) {
  try {
    const r = await fetch(url, { redirect: noFollow ? 'manual' : 'follow' });
    return {
      status: r.status,
      location: r.headers.get('location'),
      contentType: r.headers.get('content-type'),
      body: await r.text(),
    };
  } catch (e) { return { status: 0, body: '', error: String(e && e.message || e) }; }
}

/* ══ self-test（判定そのものを、わざと壊して赤にする） ══════════════════ */
if (process.argv.includes('--self-test')) {
  let pass = 0, fail = 0;
  const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + e.message); } };
  console.log('\n[check-deployed-version --self-test] 判定そのものが空振りしていないか');
  T('同じ刻印なら緑', () => { if (!judgeVersion('abc12345', ['abc12345']).ok) throw new Error('緑にならない'); });
  T('★配信が古い刻印なら赤', () => { if (judgeVersion('abc12345', ['0000ffff']).ok) throw new Error('赤にならない'); });
  T('★刻印が1つも無ければ赤', () => { if (judgeVersion('abc12345', []).ok) throw new Error('赤にならない'); });
  T('★1つでも違う刻印が混ざれば赤（貼り忘れ）', () => { if (judgeVersion('abc12345', ['abc12345', '0000ffff']).ok) throw new Error('赤にならない'); });
  T('★遅れて読む部品を拾える（_loadScript）', () => {
    const a = assetsOf("<script src=\"js/a.js?v=1\"></script>x _loadScript('lib/table-refs.js' + v)");
    if (!a.includes('lib/table-refs.js')) throw new Error('拾えない: ' + a.join(','));
    if (!a.includes('js/a.js')) throw new Error('静的な物を拾えない: ' + a.join(','));
  });
  T('★外のCDNは配信の検査に混ぜない', () => {
    const a = assetsOf('<script src="https://cdn.example/x.js"></script>');
    if (a.length) throw new Error('混ざった: ' + a.join(','));
  });
  T('刻印を取り出せる', () => {
    const s = stampsOf('<script src="js/a.js?v=af48d21f"></script>');
    if (s.join() !== 'af48d21f') throw new Error('取れない: ' + s.join());
  });
  T('★このリポジトリが養う配信を決められている', () => {
    if (!/^https:\/\//.test(DEFAULT_HOST)) throw new Error('配信の住所が決まっていない');
  });
  /* ★2026-08-18 自分で踏んだ穴。redirect を追って status だけ見ると、
     .js が1本も無い配信でも「12本ぜんぶ OK」になった。 */
  const SIG = { name: 'lib/table-refs.js', head: '/* table-refs.js — ' };
  T('中身の署名が入っていれば緑', () => {
    const v = judgeAsset({ status: 200, contentType: 'application/javascript', body: '/* table-refs.js — ★…' }, SIG);
    if (v.state !== 'ok') throw new Error('緑にならない: ' + v.why);
  });
  T('★保護のログイン画面へ飛ばされたら 緑にも赤にもせず 🟡未測定', () => {
    const v = judgeAsset({ status: 302, location: 'https://vercel.com/sso-api?url=…', body: 'Redirecting...' }, SIG);
    if (v.state !== 'unknown') throw new Error('未測定にならない: ' + v.state);
  });
  T('★200でも中身が保護の画面なら 🟡未測定（嘘の緑を作らない）', () => {
    const v = judgeAsset({ status: 200, contentType: 'text/html', body: 'Redirecting... to _vercel/sso' }, SIG);
    if (v.state !== 'unknown') throw new Error('未測定にならない: ' + v.state);
  });
  T('★200でも .js の中身がHTMLなら赤', () => {
    const v = judgeAsset({ status: 200, contentType: 'text/html; charset=utf-8', body: '<!doctype html><html>…' }, SIG);
    if (v.state !== 'ng') throw new Error('赤にならない: ' + v.state);
  });
  T('★200でも中身がうちのファイルと違えば赤（署名が無い）', () => {
    const v = judgeAsset({ status: 200, contentType: 'application/javascript', body: 'console.log("別の物")' }, SIG);
    if (v.state !== 'ng') throw new Error('赤にならない: ' + v.state);
  });
  T('★404は赤', () => {
    if (judgeAsset({ status: 404, body: '' }, SIG).state !== 'ng') throw new Error('赤にならない');
  });
  T('★繋がらなければ赤（未測定に逃がさない）', () => {
    if (judgeAsset({ status: 0, body: '', error: 'ENOTFOUND' }, SIG).state !== 'ng') throw new Error('赤にならない');
  });
  T('★保護でない転送は赤（別の所へ飛ばされている）', () => {
    const v = judgeAsset({ status: 308, location: 'https://other.example/', body: '' }, SIG);
    if (v.state !== 'ng') throw new Error('赤にならない: ' + v.state);
  });
  T('★改行の違い(CRLF/LF)だけで赤にしない（比べる前にそろえる）', () => {
    const v = judgeAsset({ status: 200, contentType: 'application/javascript', body: '/* a\r\nb */' },
      { name: 'x.js', head: '/* a\nb */' });
    if (v.state !== 'ok') throw new Error('嘘の赤が出た: ' + v.why);
  });
  console.log('\n  ── 実測 ── わざと壊した通り数を含め ' + (pass + fail) + ' 件 / 赤にできた ' + pass + ' 件');
  console.log('\n' + pass + ' passed, ' + fail + ' failed');
  process.exitCode = fail ? 1 : 0;
} else {
  /* ══ 本番（実配信を叩く。★週1＋手動。押すたびに叩かない★） ══════════ */
  const hostArg = process.argv.indexOf('--host');
  const HOST = (hostArg > 0 && process.argv[hostArg + 1]) ? process.argv[hostArg + 1].replace(/\/$/, '') : DEFAULT_HOST;
  const JSON_OUT = process.argv.includes('--json');
  const want = buildHash(ROOT);

  const pages = [];
  for (const p of PAGES) {
    if (!fs.existsSync(path.join(ROOT, p))) { pages.push({ page: p, state: 'ng', why: '★このリポジトリに ' + p + ' が無い' }); continue; }
    const r = await get(HOST + '/' + p, true);
    const a = judgeAsset(r, { name: p, head: '' });
    if (a.state !== 'ok') { pages.push({ page: p, status: r.status, ...a }); continue; }
    const stamps = stampsOf(r.body);
    const v = judgeVersion(want, stamps);
    pages.push({ page: p, status: 200, stamps, state: v.ok ? 'ok' : 'ng', why: v.why });
  }

  /* 部品が配信に実在するか（★遅れて読む物を含める／中身の署名まで見る★） */
  const localHtml = fs.existsSync(path.join(ROOT, ASSET_PAGE)) ? fs.readFileSync(path.join(ROOT, ASSET_PAGE), 'utf8') : '';
  const assets = assetsOf(localHtml);
  const assetResults = [];
  for (const a of assets) {
    /* ★署名＝うちのファイルの頭。これが返ってこなければ「別の物」★
       先頭の空白/BOMを避けて、実在する文字列を60字だけ使う */
    let sigHead = '';
    try {
      const local = fs.readFileSync(path.join(ROOT, a), 'utf8');
      sigHead = local.replace(/^﻿/, '').trim().slice(0, 60);
    } catch (e) { sigHead = ''; }
    const r = await get(HOST + '/' + a, true);
    assetResults.push({ asset: a, status: r.status, ...judgeAsset(r, { name: a, head: sigHead }) });
  }

  const ngPages = pages.filter(x => x.state === 'ng');
  const unkPages = pages.filter(x => x.state === 'unknown');
  const ngAssets = assetResults.filter(x => x.state === 'ng');
  const unkAssets = assetResults.filter(x => x.state === 'unknown');
  const mark = s => (s === 'ok' ? '✓' : s === 'unknown' ? '🟡' : '✗');

  if (JSON_OUT) {
    console.log(JSON.stringify({ host: HOST, want, ngPages: ngPages.length, unkPages: unkPages.length, ngAssets: ngAssets.length, unkAssets: unkAssets.length, pages, assets: assetResults }, null, 1));
  } else {
    console.log('\n[check-deployed-version] 配信されている物が 今のコードと同じ版か');
    console.log('  配信先      ' + HOST + (IS_STAGING ? '（テスト線）' : '（本番）'));
    console.log('  今のコード  ?v=' + want + '\n');
    console.log('■ 画面');
    pages.forEach(x => console.log('  ' + mark(x.state) + ' ' + x.page
      + (x.stamps ? '  配信=?v=' + x.stamps.join(',') : '') + (x.why ? '\n      ' + x.why : '')));
    console.log('\n■ ' + ASSET_PAGE + ' が読む部品（★遅れて読む物を含む／中身の署名まで見る★）');
    assetResults.forEach(x => console.log('  ' + mark(x.state) + ' ' + String(x.status).padEnd(4) + ' ' + x.asset + (x.why ? '  ／ ' + x.why : '')));
    console.log('\n── 実測 ──');
    console.log('  画面 ' + pages.length + '枚を数えて OK ' + (pages.length - ngPages.length - unkPages.length) + ' / NG ' + ngPages.length + ' / 🟡未測定 ' + unkPages.length);
    console.log('  部品 ' + assetResults.length + '本を数えて OK ' + (assetResults.length - ngAssets.length - unkAssets.length) + ' / NG ' + ngAssets.length + ' / 🟡未測定 ' + unkAssets.length);
    if (unkPages.length || unkAssets.length) {
      console.log('\n🟡 ★保護がかかっていて中身を読めない＝「異常なし」ではありません（未測定）★');
      console.log('  保護された配信は ★どのパスでも 302→ログイン画面 200★ を返すので、');
      console.log('  status だけ見ると「部品が全部そろっている」という嘘の緑になります（2026-08-18 に踏んだ）。');
    }
    if (ngPages.length || ngAssets.length) {
      console.log('\n★直し方★ 配信の合図が届いていない事がある（2026-08-18 に実際に起きた）。');
      console.log('  ・意味のある1コミットを push し直す（★空コミットは使わない＝CIが動かない★）');
      console.log('  ・または配信の画面で Redeploy');
      console.log('  ・押した後は ★間を空けて1回だけ★ 確かめる（叩き続けない）');
    }
  }
  /* ★未測定も緑にしない★（保護で読めないのを「異常なし」と書かない） */
  if (ngPages.length || ngAssets.length) process.exitCode = 3;
  else if (unkPages.length || unkAssets.length) process.exitCode = 2;
}
