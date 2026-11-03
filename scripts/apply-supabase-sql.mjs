/* apply-supabase-sql.mjs — ★倉庫に SQL を 当てる（門番つき）★ 2026-08-29
 *
 *  ★形は 飲み屋・ダイコメの道具から 借りた★（借りてよいのは 道具・測り方・試験）。
 *
 *  ★安全弁（1つでも 引っかかったら 1文字も 当てない）★
 *    ①門番 scripts/sql-guard.mjs を 通す（足すだけ・締めるだけ・exally. の部屋だけ）
 *    ②★向き先は repo の js/supa-config.js から 読む★＝この repo の倉庫以外には 当たらない
 *    ③本番へ当てる時は ★--honban★ を 明示する（既定は テスト倉庫）
 *    ④当てた後に ★効いたかを 数える★（棚・列・RLS・決まり・権限）＝「書いた」≠「効いた」
 *
 *  使い方:
 *    node scripts/apply-supabase-sql.mjs --probe <file...>     … 門番と 向き先だけ 見る（当てない）
 *    node scripts/apply-supabase-sql.mjs <file...>             … テスト倉庫へ 当てる
 *    node scripts/apply-supabase-sql.mjs --honban <file...>    … ★本番の倉庫へ 当てる★
 *
 *  鍵: %TEMP%/nomiya-db-url-prod.json（本番） / %TEMP%/nomiya-db-url.json（テスト）
 *      ★中身は 1文字も 画面に 出さない★
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { guard } from './sql-guard.mjs';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const 引数 = process.argv.slice(2);
const 下見 = 引数.includes('--probe');
const 本番 = 引数.includes('--honban');
const ファイル達 = 引数.filter((a) => !a.startsWith('--'));

function 止める(なぜ) {
  console.error('\n★中止★ ' + なぜ);
  console.error('★1文字も 当てていません★');
  process.exit(1);
}

/* ── ② 向き先は repo から 読む ─────────────────────────────
   ★倉庫の番号を ここに 直接 書かない★（見張り tests/no-hardcoded-supa.test.mjs が 赤にする）。
   理由＝本番repo から 作った複製を テストrepo に 持ってくると 直書きが 付いてきて
   ★テストのつもりで 本番の倉庫を 触る★という 最悪の事故になる（tests/repo-supa.mjs の頭より）。
   ⇒ ★向き先は repoSupa()★ ／ ★本番/テストの見分けは check-warehouse-pointers.mjs の定数★
      （どちらも repo に 1か所しか 無い物を 借りる。写さない） */
const { repoSupa } = await import('../tests/repo-supa.mjs');
const { PROD_REF, TEST_REF } = await import('./check-warehouse-pointers.mjs');
let REF;
try { REF = repoSupa(ROOT).ref; } catch (e) { 止める('向き先を 読めない: ' + e.message); }
if (!REF) 止める('js/supa-config.js から 倉庫を 読めない');

const 実際は = REF === PROD_REF ? '本番' : REF === TEST_REF ? 'テスト' : '知らない倉庫';
if (実際は === '知らない倉庫') 止める('この repo の向き先 ' + REF + ' は 本番でも テストでもない');
if (本番 && 実際は !== '本番') 止める('--honban と 言われたが この repo は ' + 実際は + ' を 指している');
if (!本番 && 実際は === '本番') {
  止める('この repo は ★本番★ を 指している。本番へ 当てるなら --honban を 付ける');
}

console.log('[apply-supabase-sql] 倉庫 = ' + REF + '（★' + 実際は + '★）');
console.log('  当てるファイル … ' + (ファイル達.join(', ') || '（無し）'));
if (!ファイル達.length) 止める('ファイルが 指定されていない');

/* ── ① 門番 ───────────────────────────────────────── */
const 中身 = [];
let 落ちた = 0;
for (const f of ファイル達) {
  const p = path.isAbsolute(f) ? f : path.join(ROOT, f);
  if (!fs.existsSync(p)) 止める('ファイルが 無い: ' + f);
  const sql = fs.readFileSync(p, 'utf8');
  const g = guard(sql, { prefix: 'exally.' });
  console.log('\n  ── ' + f + ' ──');
  console.log('     触る棚 … ' + (g.tables.join(', ') || '（無し）'));
  console.log('     do $$ の塊 … ' + g.doBlocks + '個');
  if (!g.ok) {
    落ちた++;
    for (const r of g.reasons) console.log('     ★止めた★ ' + r);
  } else {
    console.log('     ✓ 門番を 通った');
  }
  中身.push({ f, sql, tables: g.tables });
}
if (落ちた) 止める('門番に ' + 落ちた + '本 落ちた（★1本でも 落ちたら 1文字も 当てない★）');

if (下見) {
  console.log('\n★--probe＝ここまで。1文字も 当てていません★');
  process.exit(0);
}

/* ── 鍵 ───────────────────────────────────────────── */
const 鍵の場所 = path.join(os.tmpdir(), 実際は === '本番' ? 'nomiya-db-url-prod.json' : 'nomiya-db-url.json');
if (!fs.existsSync(鍵の場所)) 止める('鍵が 無い（' + path.basename(鍵の場所) + '）＝司さんに 作り直しを 頼む');
let TOKEN;
try { TOKEN = JSON.parse(fs.readFileSync(鍵の場所, 'utf8')).token; } catch (e) { 止める('鍵を 読めない'); }
if (!TOKEN || TOKEN.indexOf('sbp_') !== 0) 止める('鍵の形が おかしい');

async function 投げる(sql) {
  const r = await fetch('https://api.supabase.com/v1/projects/' + REF + '/database/query', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + TOKEN,
      'Content-Type': 'application/json',
      'User-Agent': 'claude-code-exally/1.0',   // ★付けないと Cloudflare が 403 で弾く★
    },
    body: JSON.stringify({ query: sql }),
  });
  const t = await r.text();
  if (!r.ok) throw new Error('HTTP ' + r.status + ' ' + t.slice(0, 300));
  try { return JSON.parse(t); } catch (e) { return t; }
}

/* ── 当てる ───────────────────────────────────────── */
console.log('\n★当てます★（' + 実際は + ' / ' + REF + '）');
for (const { f, sql } of 中身) {
  try {
    await 投げる(sql);
    console.log('  ✓ 当てた … ' + f);
  } catch (e) {
    止める('当てている途中で 失敗: ' + f + ' … ' + e.message);
  }
}

/* ── ④ 効いたかを 数える（「書いた」≠「効いた」）───────────── */
const 棚達 = Array.from(new Set(中身.flatMap((x) => x.tables)));
console.log('\n★効いたかを 数える★');
for (const t of 棚達) {
  const [部屋, 棚] = t.split('.');
  const q = async (sql) => 投げる(sql);
  const 列 = await q("select count(*) c from information_schema.columns where table_schema='" + 部屋 + "' and table_name='" + 棚 + "'");
  const rls = await q("select relrowsecurity r from pg_class where oid = '" + t + "'::regclass");
  const 決まり = await q("select polname, polcmd from pg_policy where polrelid = '" + t + "'::regclass order by polname");
  const 権限 = await q("select grantee, string_agg(privilege_type, ',' order by privilege_type) p from information_schema.role_table_grants where table_schema='" + 部屋 + "' and table_name='" + 棚 + "' group by grantee order by grantee");
  const 行 = await q('select count(*) c from ' + t);
  console.log('  ' + t);
  console.log('    列 … ' + 列[0].c + '本 ／ RLS … ' + (rls[0].r ? '★入っている★' : '★入っていない★'));
  console.log('    決まり … ' + (決まり.map((x) => x.polname + '(' + x.polcmd + ')').join(' / ') || '★0本★'));
  for (const g of 権限) console.log('    権限 … ' + g.grantee + ' = ' + g.p);
  console.log('    行数 … ' + 行[0].c + '行');
  if (!rls[0].r) 止める('★RLS が 効いていない★: ' + t);
  if (!決まり.length) 止める('★決まりが 1本も 無い★: ' + t);
  const 誰でも = 権限.filter((g) => g.grantee === 'anon' || g.grantee === 'PUBLIC');
  if (誰でも.length) 止める('★anon/PUBLIC に 権限が 残っている★: ' + 誰でも.map((g) => g.grantee + '=' + g.p).join(' '));
}
console.log('\nAPPLY RESULT: OK');
