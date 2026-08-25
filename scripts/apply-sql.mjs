/* apply-sql.mjs — ★倉庫へ SQL を当てる（門番つき）★
 *
 *  ★決まり（他アプリと同じ流儀・飲み屋/ダイコメで先に作った物に合わせる）★
 *    ①★向き先は この repo の js/supa-config.js から読む★＝repoの向き先以外には 当たらない
 *    ②★本番の倉庫には 当てない★（--i-know-this-is-prod を付けない限り 止まる）
 *    ③★消す系（drop / truncate / delete / update）が1つでも混ざったら 1文字も当てない★
 *       （★90日の掃除は 別の口（--souji）★＝当てる物と 消す物を 混ぜない）
 *    ④★部屋は exally だけ★（他アプリの部屋を触らない）
 *    ⑤★当てた後に 表と列と権限を数えて 出す★（「書いた」で終わらせない）
 *
 *  使い方:
 *    node scripts/apply-sql.mjs supabase/xxx.sql        … 当てる
 *    node scripts/apply-sql.mjs --check supabase/xxx.sql … 門番だけ通す（1文字も当てない）
 *    node scripts/apply-sql.mjs --self-test             … ★門番が本当に止めるかを 数える★
 *    node scripts/apply-sql.mjs --sql "select …"        … 読むだけ（select で始まる物だけ）
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.join(__dirname, '..');
/* ★倉庫の名前を ここに書かない★（見張り tests/no-hardcoded-supa.test.mjs が捕まえた・2026-08-26）
   ＝向き先を持っているのは ★js/supa-config.js だけ★ という約束。
   ⇒ 本番かどうかは ★その repo の supa-config.js の env★ で決める。 */

/* ── ★門番★（純関数＝わざと危ない物を食わせて 試験できる）── */
export function 門番(sql, opt) {
  opt = opt || {};
  const 中 = String(sql || '');
  const 悪い = [];
  /* 注記(コメント)は 見ない＝注記に書いた言葉で 止まらないように */
  const 本文 = 中.replace(/--[^\n]*/g, ' ').replace(/\/\*[\s\S]*?\*\//g, ' ');
  if (!本文.trim()) 悪い.push('中身が空');
  const 消す系 = [
    [/\bdrop\s+(table|schema|database|index|view|function)\b/i, 'drop'],
    [/\btruncate\b/i, 'truncate'],
    [/\bdelete\s+from\b/i, 'delete'],
    [/\bupdate\s+[a-z_."]+\s+set\b/i, 'update'],
    [/\balter\s+table\s+[a-z_."]+\s+drop\b/i, 'alter…drop'],
  ];
  for (const [re, 名] of 消す系) if (re.test(本文)) 悪い.push('消す系が混ざっている: ' + 名);
  /* ★部屋は exally だけ★ */
  const 部屋 = new Set();
  const re = /\b(?:create|alter|comment on)\s+(?:table|index|view|materialized view)?\s*(?:if not exists\s+)?([a-z_][a-z0-9_]*)\./gi;
  let m;
  while ((m = re.exec(本文))) 部屋.add(m[1].toLowerCase());
  for (const p of 部屋) if (p !== 'exally') 悪い.push('別の部屋を触っている: ' + p);
  /* ★書く物が1つも無いのに「当てる」と言っていないか★ */
  if (!/\b(create|alter|grant|revoke|comment)\b/i.test(本文) && !opt.読むだけ) 悪い.push('当てる物が1つも無い');
  return { ok: 悪い.length === 0, 悪い: 悪い, 部屋: [...部屋] };
}

export function 向き先(repoRoot) {
  const src = fs.readFileSync(path.join(repoRoot, 'js/supa-config.js'), 'utf8');
  const m = /https:\/\/([a-z0-9]+)\.supabase\.co/.exec(src);
  if (!m) throw new Error('★向き先が 読めない（js/supa-config.js）★');
  return m[1];
}
/** ★この repo が 本番向きか★（★フォルダ名や repo名では決めない★＝env を読む） */
export function 本番か(repoRoot) {
  /* ★注記(コメント)を先に外す★＝注記に書いた「env:'prod'」を 本物と読み違えた（2026-08-26・自分で踏んだ） */
  const src = fs.readFileSync(path.join(repoRoot, 'js/supa-config.js'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, ' ').replace(new RegExp('//[^' + String.fromCharCode(10) + ']*', 'g'), ' ');
  const m = /env\s*:\s*'(prod|test)'/.exec(src);
  if (!m) throw new Error('★env が 読めない＝本番かどうか 分からないので 止める★');
  return m[1] === 'prod';
}

function 鍵() {
  for (const p of [path.join(os.tmpdir(), 'nomiya-db-url.json'), path.join(os.tmpdir(), 'nomiya-db-url-prod.json')]) {
    try {
      const t = JSON.parse(fs.readFileSync(p, 'utf8')).token;
      if (t) return t;
    } catch (e) { /* 無ければ次 */ }
  }
  throw new Error('★鍵が無い★ %TEMP%\\nomiya-db-url.json（司さんに作り直しを頼む）');
}

export async function 当てる(ref, sql) {
  const res = await fetch('https://api.supabase.com/v1/projects/' + ref + '/database/query', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + 鍵(),
      'Content-Type': 'application/json',
      'User-Agent': 'exally-claude/1.0',   /* ★これが無いと Cloudflare が 403 で弾く★ */
    },
    body: JSON.stringify({ query: sql }),
  });
  const t = await res.text();
  if (!res.ok) throw new Error('★当たらなかった★ ' + res.status + ' ' + t.slice(0, 400));
  try { return JSON.parse(t); } catch (e) { return t; }
}

/* ── 走らせる ── */
const args = process.argv.slice(2);
if (args.includes('--self-test')) {
  const 表 = [
    ['まともな物', 'create table if not exists exally.a (id int);', true],
    ['★drop が混ざる★', 'create table exally.a (id int); drop table exally.b;', false],
    ['★truncate が混ざる★', 'truncate exally.a;', false],
    ['★delete が混ざる★', 'delete from exally.a where 1=1;', false],
    ['★update が混ざる★', 'update exally.a set id = 1;', false],
    ['★別の部屋を触る★', 'create table public.a (id int);', false],
    ['★別アプリの部屋を触る★', 'create table nomiya.a (id int);', false],
    ['★中身が空★', '   ', false],
    ['★注記に書いた drop では 止まらない★', '-- drop table は しない\ncreate table exally.a (id int);', true],
    ['★当てる物が1つも無い★', 'select 1;', false],
    ['権限だけ', 'revoke all on exally.a from anon;', true],
  ];
  let 緑 = 0, 赤 = 0;
  console.log('');
  console.log('[apply-sql] ★門番が 本当に止めるか★');
  for (const [名, sql, 待つ] of 表) {
    const r = 門番(sql);
    if (r.ok === 待つ) { 緑++; console.log('  ok   ' + 名); }
    else { 赤++; console.log('  NG   ' + 名 + ' … ' + JSON.stringify(r.悪い)); }
  }
  /* ★本番の倉庫を 弾くか★ */
  if (!本番か(ROOT)) { 緑++; console.log('  ok   ★この repo は 本番向きではない（env=test）★'); }
  else { 赤++; console.log('  NG   ★この repo が 本番向き（env=prod）★'); }
  /* ★倉庫の名前を この道具に 直書きしていない事★（見張りと同じ決まりを 自分でも数える） */
  const 自分 = fs.readFileSync(new URL(import.meta.url), 'utf8');
  if (!/[a-z0-9]{20}\.supabase\.co|['"][a-z]{20}['"]/.test(自分)) { 緑++; console.log('  ok   ★倉庫の名前を 直書きしていない★'); }
  else { 赤++; console.log('  NG   ★倉庫の名前を 直書きしている★'); }
  console.log('');
  console.log('  ' + 緑 + ' 緑 / ' + 赤 + ' 赤');
  process.exit(赤 ? 1 : 0);
}

/* ── ★どの口を使うか（ここから下は 1つだけ 走る）── */
const ファイル = args.filter((a) => !a.startsWith('--'));

if (args.includes('--souji')) {
  /* ★90日の掃除★（当てる物と 消す物を 混ぜない＝口を分ける）
     ★消せるのは この1文だけ★＝場所も 日数も ここに焼き付けてある（引数で変えられない）。
     ★誰が いつ 回すか★ … サーバが ★その日の最初の書き込みの時に 1回だけ★（docs/SPEC_kazoeba.md）。
     ここは ★人が手で回す口★＝初回の確認と、困った時のため。 */
  const ref0 = 向き先(ROOT);
  if (本番か(ROOT) && !args.includes('--i-know-this-is-prod')) { console.error('★本番の倉庫★'); process.exit(1); }
  const 前 = await 当てる(ref0, 'select count(*) as n from exally.ai_tsukatta');
  const 消す = await 当てる(ref0,
    "delete from exally.ai_tsukatta where oshita < now() - interval '90 days' returning id");
  const 後 = await 当てる(ref0, 'select count(*) as n from exally.ai_tsukatta');
  console.log('掃除 … 前 ' + 前[0].n + '行 → ★消した ' + 消す.length + '行★ → 後 ' + 後[0].n + '行');
} else if (args.includes('--sql')) {
  /* ★読むだけの口★ */
  const sql = args[args.indexOf('--sql') + 1];
  if (!/^\s*(select|begin|with|set)/i.test(sql)) { console.error('★読むだけの口には select しか渡せない★'); process.exit(1); }
  const ref = 向き先(ROOT);
  if (本番か(ROOT) && !args.includes('--i-know-this-is-prod')) { console.error('★本番の倉庫★'); process.exit(1); }
  console.log(JSON.stringify(await 当てる(ref, sql), null, 1));
} else if (!ファイル.length) {
  console.error('使い方: node scripts/apply-sql.mjs <file.sql>');
  process.exit(1);
} else {
  const ref = 向き先(ROOT);
  console.log('向き先 … ' + ref + (本番か(ROOT) ? '（★本番★）' : '（DB-test）'));
  if (本番か(ROOT) && !args.includes('--i-know-this-is-prod')) {
    console.error('★本番の倉庫には 当てません（指示役に声をかけてから）★');
    process.exit(1);
  }
  let 全部 = '';
  let 止まった = false;
  for (const f of ファイル) {
    const sql = fs.readFileSync(path.join(ROOT, f), 'utf8');
    const r = 門番(sql);
    console.log((r.ok ? '  通した  ' : '  ★止めた★  ') + f + (r.ok ? '' : ' … ' + r.悪い.join(' / ')));
    if (!r.ok) 止まった = true;
    全部 += sql + String.fromCharCode(10);
  }
  if (止まった) { console.error('★1本でも門番に落ちたら 1文字も当てない★'); process.exit(1); }
  if (args.includes('--check')) console.log('★--check なので 1文字も当てていません★');
  else {
    const out = await 当てる(ref, 全部);
    console.log('当てました:', JSON.stringify(out).slice(0, 200));
  }
}
/* ★fetch の後に process.exit すると Windows の libuv が悲鳴を上げる★＝静かに終わる */
