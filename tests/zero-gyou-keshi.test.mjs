/*
 * ★消しが 0行でも「削除しました」と 出た（2026-11-14）★
 *   鍵が 切れて anon で 送る・他の口の 行・もう 無い 行は、倉庫が 0行で 204・error null を 返す。
 *   前の js/suite-data.js の partners.remove・ledger.remove は error だけを 見て ok:true を 返していた。
 *   直し＝.select('id') で 書けた 行を 数え、0 なら not-found／ログインが 無ければ no-user。
 *   作り物（tests/fake-supa.js）も 本物と 同じく、select を 付けない update は data:null を 返す。
 *
 *  使い方: node tests/zero-gyou-keshi.test.mjs ／ --self-test＝守りを 1本ずつ 抜いて 回し、どれも 赤に なるか
 */
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { createRequire } from 'node:module';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const require = createRequire(import.meta.url);
const HERE = path.dirname(fileURLToPath(import.meta.url));
const SRC = path.join(HERE, '..', 'js', 'suite-data.js');
const { createFakeSupa } = require('./fake-supa.js');

/* ★守り（js/suite-data.js の 字）＝--self-test=<名> で 1本だけ 抜く★ */
const 守り = {
  '書けた行を数える': [
    "return n > 0 ? { ok: true } : { ok: false, reason: 'not-found' };",
    'return { ok: true };'
  ],
  '取引先の消しの select': [
    "update({ deleted_at: nowIso(), updated_at: nowIso() }).eq('id', id).select('id'))\n              .then(kesetaKa);\n          }).catch(function (e) { return { ok: false, reason: err(e) }; });\n        }\n      },\n\n      /* ═══ 日次",
    "update({ deleted_at: nowIso(), updated_at: nowIso() }).eq('id', id))\n              .then(kesetaKa);\n          }).catch(function (e) { return { ok: false, reason: err(e) }; });\n        }\n      },\n\n      /* ═══ 日次"
  ],
  '記録の消しの select': [
    "from('pay_ledger').update({ deleted_at: nowIso(), updated_at: nowIso() }).eq('id', id).select('id'))",
    "from('pay_ledger').update({ deleted_at: nowIso(), updated_at: nowIso() }).eq('id', id))"
  ],
  'ログインが無ければ止める': [
    "if (!u) return { ok: false, reason: 'no-user' };   // 鍵切れ＝「ログインし直してください」\n            return Promise.resolve(sb.from('pay_partners')",
    "return Promise.resolve(sb.from('pay_partners')"
  ],
  'ログインが無ければ止める（記録）': [
    "if (!u) return { ok: false, reason: 'no-user' };   // 鍵切れ＝「ログインし直してください」\n            return Promise.resolve(sb.from('pay_ledger')",
    "return Promise.resolve(sb.from('pay_ledger')"
  ]
};

if (process.argv.includes('--self-test')) {
  let 赤 = 0;
  for (const 名 of Object.keys(守り)) {
    const r = spawnSync(process.execPath, [fileURLToPath(import.meta.url), '--self-test=' + 名], { encoding: 'utf8' });
    const ng = (r.stdout.match(/^ {2}NG /gm) || []).length;
    const 赤い = r.status !== 0 && ng > 0;
    if (赤い) 赤++;
    console.log('  ' + (赤い ? 'ok ' : 'NG ') + '  守り「' + 名 + '」を 抜く→赤 ' + ng + (r.status === 2 ? '（抜く字が無い）' : ''));
  }
  console.log('\n守り ' + Object.keys(守り).length + '本' + (赤 === Object.keys(守り).length ? ' 全部 赤' : ' のうち 赤は ' + 赤 + '本＝空振りが在る'));
  process.exit(赤 === Object.keys(守り).length ? 0 : 1);
}

const 抜く名 = (process.argv.find((a) => a.startsWith('--self-test=')) || '').slice('--self-test='.length);
let SuiteData;
if (抜く名) {
  const [前, 後] = 守り[抜く名] || [];
  const 字 = fs.readFileSync(SRC, 'utf8');
  if (!前 || 字.split(前).length !== 2) { console.log('  NG   抜く字が 1か所に 見つからない: ' + 抜く名); process.exit(2); }
  const 写し = path.join(fs.mkdtempSync(path.join(os.tmpdir(), 'zg-self-')), 'suite-data.js');
  fs.writeFileSync(写し, 字.replace(前, 後));
  SuiteData = require(写し);
} else {
  SuiteData = require(SRC);
}

function setup() {
  const tables = {
    pay_employees: [], pay_companies: [], pay_org: [], exally_entitlements: [],
    pay_partners: [
      { id: 'pt_mine', account_id: 'u1', sort: 0, data: { name: '自分の取引先' }, deleted_at: null, updated_at: '2026-07-01T00:00:00Z' },
      { id: 'pt_tanin', account_id: 'u2', sort: 0, data: { name: '他の口の取引先' }, deleted_at: null, updated_at: '2026-07-01T00:00:00Z' }
    ],
    pay_ledger: [
      { id: 'lg_mine', account_id: 'u1', employee_id: 'e1', ymd: '2026-07-01', data: { uriage: 100 }, deleted_at: null, updated_at: '2026-07-01T00:00:00Z' },
      { id: 'lg_tanin', account_id: 'u2', employee_id: 'x9', ymd: '2026-07-01', data: { uriage: 999 }, deleted_at: null, updated_at: '2026-07-01T00:00:00Z' }
    ]
  };
  const sb = createFakeSupa({ uid: 'u1', tables, pk: { pay_companies: 'account_id', pay_org: 'account_id', exally_entitlements: ['account_id', 'app'] } });
  const sd = SuiteData.create({ client: sb });
  return { sb, sd, tables };
}

const T = [];
function test(name, fn) { T.push({ name, fn }); }
function eq(a, b, m) { if (a !== b) throw new Error((m || '') + ' 期待=' + JSON.stringify(b) + ' 実際=' + JSON.stringify(a)); }

const 口 = [['取引先', 'partners', 'pay_partners', 'pt'], ['記録', 'ledger', 'pay_ledger', 'lg']];
for (const [名, api, 表, 頭] of 口) {
  test(名 + '：自分の行を消す→ok・deleted_at が入る', async () => {
    const { sd, tables } = setup();
    const r = await sd[api].remove(頭 + '_mine');
    eq(r.ok, true, 'ok');
    eq(!!tables[表].find((x) => x.id === 頭 + '_mine').deleted_at, true, 'deleted_at');
  });
  test(名 + '：他の口の行（RLS で 0行）→not-found・行は消えない', async () => {
    const { sd, tables } = setup();
    const r = await sd[api].remove(頭 + '_tanin');
    eq(r.ok, false, 'ok');
    eq(r.reason, 'not-found', 'reason');
    eq(tables[表].find((x) => x.id === 頭 + '_tanin').deleted_at, null, 'deleted_at');
  });
  test(名 + '：無い id→not-found', async () => {
    const { sd } = setup();
    const r = await sd[api].remove(頭 + '_nai');
    eq(r.ok, false, 'ok');
    eq(r.reason, 'not-found', 'reason');
  });
  test(名 + '：ログインが無い（鍵切れ）→no-user・倉庫へ書かない', async () => {
    const { sb, sd, tables } = setup();
    sb._uid(null);
    sb._reset();
    const r = await sd[api].remove(頭 + '_mine');
    eq(r.ok, false, 'ok');
    eq(r.reason, 'no-user', 'reason');
    eq(sb._calls.filter((c) => c.op === 'update').length, 0, 'update を 送った');
    eq(tables[表].find((x) => x.id === 頭 + '_mine').deleted_at, null, 'deleted_at');
  });
  test(名 + '：倉庫の失敗→ok にしない', async () => {
    const { sb, sd } = setup();
    sb._failNext(表, 'update');
    const r = await sd[api].remove(頭 + '_mine');
    eq(r.ok, false, 'ok');
    eq(r.reason, 'fake-failure', 'reason');
  });
}

test('作り物：select を 付けない update は data:null（本物の 204 と 同じ）・付けると 行', async () => {
  const { sb } = setup();
  const a = await sb.from('pay_partners').update({ updated_at: 'x' }).eq('id', 'pt_mine');
  eq(a.data, null, 'select 無し');
  const b = await sb.from('pay_partners').update({ updated_at: 'y' }).eq('id', 'pt_mine').select('id');
  eq(Array.isArray(b.data) && b.data.length, 1, 'select 有り');
  const c = await sb.from('pay_partners').update({ updated_at: 'z' }).eq('id', 'pt_mine').select();
  eq(Array.isArray(c.data) && c.data.length, 1, '引数無しの select');
});

test('字の見張り：deleted_at を 書く update は 全部 .select( を 付けて kesetaKa で 判じる', () => {
  const 字 = fs.readFileSync(SRC, 'utf8');
  const 拾う = /\.update\(\{[^}\n]*\bdeleted_at\b[^\n]*/g;   // ★キーの 順に 頼らない（taiketsu 任意4）
  eq(('x.update({ updated_at: 1, deleted_at: 2 }).eq(1)'.match(拾う) || []).length, 1, '順を 変えた 形を 拾う');
  const 消し = 字.match(拾う) || [];
  eq(消し.length, 2, '消しの数（増えたら この試験も 見直す）');
  消し.forEach((l) => { if (!/\.select\(/.test(l)) throw new Error('select が 無い: ' + l.trim()); });
  eq((字.match(/\.then\(kesetaKa\)/g) || []).length, 2, 'kesetaKa で 判じる 数');
});

let ng = 0;
for (const t of T) {
  try { await t.fn(); console.log('  ok   ' + t.name); }
  catch (e) { ng++; console.log('  NG   ' + t.name + '\n       ' + (e && e.message)); }
}
console.log('\n' + (T.length - ng) + ' ok / ' + ng + ' NG');
process.exit(ng ? 1 : 0);
