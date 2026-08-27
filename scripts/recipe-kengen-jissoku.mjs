/* recipe-kengen-jissoku.mjs — ★覚えた手順の置き場の権限を「行数」で 実測する★
 * =============================================================================
 * ★決まり（うちが 何度も 踏んだ所）★
 *   ・★書いた≠効いた★＝SQLに書いた事ではなく ★実際に押して 行数を数える★
 *   ・★anon 0★／★本人の1行だけ★／★なりすまし insert が 弾かれる★
 *   ・★DB-test だけ★（向き先は js/supa-config.js から読む＝この道具は 名前を持たない）
 *   ・★begin … rollback★＝倉庫に 何も残さない
 *
 * 使い方: node scripts/recipe-kengen-jissoku.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { fileURLToPath } from 'node:url';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');

/* ★向き先は この repo の js/supa-config.js だけ★
   ★注記(コメント)を先に外す★＝注記の中の「env:'prod'」を 拾って
   「この repo は prod です」と 言って止まった（2026-08-27 実際に踏んだ）。
   外す道具は 借り物の正本 scripts/lib/chuki.mjs（自分で書かない）。 */
const { 注記を外す } = await import('./lib/chuki.mjs');
const conf = 注記を外す(fs.readFileSync(path.join(ROOT, 'js/supa-config.js'), 'utf8'));
const url = (conf.match(/url:\s*'([^']+)'/) || [])[1];
const key = (conf.match(/key:\s*'([^']+)'/) || [])[1];
if (!url || !key) throw new Error('★向き先が 読めない★');
const ref = url.replace('https://', '').split('.')[0];
/* ★倉庫の名前を この道具に 書かない★（見張り tests/no-hardcoded-supa.test.mjs）。
   ★repo 自身の env で 決める★＝'test' でなければ 1行も 触らない。 */
const env = (conf.match(/env:\s*'([^']+)'/) || [])[1] || 'prod';
if (env !== 'test') throw new Error('★この repo は ' + env + ' です。この道具は テストの倉庫だけ★');

let token = null;
for (const p of [path.join(os.tmpdir(), 'nomiya-db-url.json')]) {
  if (fs.existsSync(p)) {
    const j = JSON.parse(fs.readFileSync(p, 'utf8'));
    token = j.token || j.access_token || j.SUPABASE_ACCESS_TOKEN || null;
  }
}
if (!token) throw new Error('★鍵が無い★ %TEMP%\\nomiya-db-url.json');

async function q(sql) {
  const r = await fetch('https://api.supabase.com/v1/projects/' + ref + '/database/query', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
      'User-Agent': 'exally-kengen-jissoku',
    },
    body: JSON.stringify({ query: sql }),
  });
  const t = await r.text();
  if (!r.ok) throw new Error(r.status + ' ' + t.slice(0, 300));
  try { return JSON.parse(t); } catch (e) { return t; }
}

const 出 = [];
const 言う = (s) => { 出.push(s); console.log(s); };

console.log('');
console.log('[recipe 権限の実測] 向き先 … ' + ref + '（DB-test）');

/* ① 権限（誰に 何が 渡っているか） */
const g = await q(`select grantee, string_agg(privilege_type, ',' order by privilege_type) k
  from information_schema.role_table_grants
  where table_schema='exally' and table_name='recipe' group by grantee order by 1`);
言う('  権限 … ' + g.map((x) => x.grantee + '=' + x.k).join(' / '));
const anon行 = g.filter((x) => x.grantee === 'anon');
言う('  ★anon に渡っている権限 … ' + anon行.length + '件★' + (anon行.length ? ' ← ★駄目★' : ''));

/* ② RLS が 入っているか＋ポリシーの数 */
const p1 = await q(`select relrowsecurity, relforcerowsecurity from pg_class where oid='exally.recipe'::regclass`);
言う('  RLS … ' + (p1[0].relrowsecurity ? '入っている' : '★入っていない★'));
const pol = await q(`select polname, pg_get_expr(polqual,polrelid) q, pg_get_expr(polwithcheck,polrelid) w
  from pg_policy where polrelid='exally.recipe'::regclass order by polname`);
言う('  ポリシー … ' + pol.length + '本（' + pol.map((x) => x.polname).join(' / ') + '）');

/* ③ ★行数で 数える★（begin … rollback＝倉庫に 何も残さない） */
const A = '11111111-1111-1111-1111-111111111111';
const B = '22222222-2222-2222-2222-222222222222';
const 数える = async (誰, sql) => {
  const r = await q(`begin;
    insert into exally.recipe (mochinushi, na, tanomi, shimon, yoyaku, tejun)
      values ('${A}','Aの手順','a','x','{}','[]'), ('${B}','Bの手順','b','y','{}','[]');
    set local role authenticated;
    set local request.jwt.claims = '{"sub":"${誰}","role":"authenticated"}';
    ${sql}
  rollback;`);
  return r;
};

const 見えた = await 数える(A, `select count(*)::int as n from exally.recipe;`);
/* ★返事は「文ごとの結果の並び」★＝最後の select の 1行目を 取る（形を 決めつけない） */
function 数を取る(r) {
  if (!Array.isArray(r)) return null;
  for (let i = r.length - 1; i >= 0; i--) {
    const x = r[i];
    if (x && typeof x === 'object' && !Array.isArray(x) && typeof x.n === 'number') return x.n;
    if (Array.isArray(x) && x[0] && typeof x[0].n === 'number') return x[0].n;
  }
  return null;
}
const 見える行 = 数を取る(見えた);
言う('  ★Aさんに 見える行 … ' + (見える行 === null ? JSON.stringify(見えた).slice(0, 120) : 見える行 + '行')
  + '★（2行 入れたうち 本人の1行だけなら 正しい）');

/* ④ なりすまし insert（Aさんが Bさんの名前で 書く）＝弾かれるか */
let なりすまし = '★通ってしまった（駄目）★';
try {
  await q(`begin;
    set local role authenticated;
    set local request.jwt.claims = '{"sub":"${A}","role":"authenticated"}';
    insert into exally.recipe (mochinushi, na, tanomi, shimon, yoyaku, tejun)
      values ('${B}','なりすまし','x','y','{}','[]');
  rollback;`);
} catch (e) {
  なりすまし = '弾かれた（' + String(e.message).slice(0, 80) + '）';
}
言う('  ★なりすまし insert … ' + なりすまし + '★');

/* ⑤ anon で 読めるか（★実際に REST を叩く★＝画面と同じ道） */
const r5 = await fetch(url + '/rest/v1/recipe?select=*&limit=5', {
  headers: { apikey: key, Authorization: 'Bearer ' + key, 'Accept-Profile': 'exally' },
});
const t5 = await r5.text();
let anon件 = '読めない';
try { const j = JSON.parse(t5); anon件 = Array.isArray(j) ? (j.length + '行') : ('駄目 ' + r5.status); } catch (e) { anon件 = '駄目 ' + r5.status; }
言う('  ★anon で 読んだ結果 … HTTP ' + r5.status + ' / ' + anon件 + '★');

/* ⑥ 残っていないか（rollback が 効いているか） */
const 残り = await q(`select count(*)::int as n from exally.recipe`);
言う('  倉庫に 残った行 … ' + 数を取る(残り) + '行（0 なら 何も残していない）');

console.log('');
/* ★anon は 401（そもそも 触れない）か 200で0行★ のどちらかなら 正しい */
const anonが駄目 = !(r5.status === 401 || r5.status === 403 || (r5.status === 200 && anon件 === '0行'));
const 駄目 = anon行.length > 0 || !p1[0].relrowsecurity || pol.length < 3
  || なりすまし.indexOf('通ってしまった') >= 0 || anonが駄目
  || 数を取る(残り) !== 0 || 見える行 !== 1;
console.log(駄目 ? '  ★どこかが 駄目です★' : '  ★全部 通りました（anon 0／本人の1行だけ／なりすまし 弾かれた／倉庫に 残さない）★');
process.exit(駄目 ? 1 : 0);
