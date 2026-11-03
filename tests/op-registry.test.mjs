/* op-registry.test.mjs — 契約の入口(lib/op-registry.js)の単体テスト。
 *   ★一番大事なのは「同じidの二重登録が【投げる】」こと。
 *     黙って上書きされると、どのエンジンで計算したか分からなくなり provenance が嘘になる。
 * 使い方: node tests/op-registry.test.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const require_ = createRequire(import.meta.url);
const R = require_(path.join(ROOT, 'lib/op-registry.js'));

let pass = 0, fail = 0;
const T = (n, fn) => { try { fn(); pass++; console.log('  ✓ ' + n); } catch (e) { fail++; console.log('  ✗ ' + n + ' — ' + (e && e.message)); } };
const threw = (fn) => { try { fn(); return false; } catch (e) { return true; } };
const op = (id) => ({ id: id, version: '1.0.0', title: 't-' + id, desc: 'd', engine: function () { return { value: 1, cells: null, warnings: [], errors: [], provenance: { op: id } }; } });

console.log('\n[op-registry] 契約の入口');

T('登録して id で引ける', () => {
  R._reset();
  R.register(op('a.b'));
  if (!R.get('a.b')) throw new Error('引けない');
  if (!R.has('a.b')) throw new Error('has が false');
});

T('無い id は null（例外にしない＝呼ぶ側が分岐できる）', () => {
  R._reset();
  if (R.get('nai') !== null) throw new Error('null でない');
});

T('★同じidの二重登録は投げる（黙って上書きしない）', () => {
  R._reset();
  R.register(op('a.b'));
  if (!threw(() => R.register(op('a.b')))) throw new Error('投げていない＝黙って上書きされる');
});

T('engine の無い物は登録できない', () => {
  R._reset();
  if (!threw(() => R.register({ id: 'x' }))) throw new Error('投げていない');
  if (!threw(() => R.register({ engine: function () { } }))) throw new Error('id無しで投げていない');
  if (!threw(() => R.register(null))) throw new Error('null で投げていない');
});

T('list はカタログを返す（id順・チャットの「何ができるか」に使う）', () => {
  R._reset();
  R.register(op('z.z')); R.register(op('a.a'));
  const l = R.list();
  if (l.length !== 2) throw new Error('件数が違う');
  if (l[0].id !== 'a.a' || l[1].id !== 'z.z') throw new Error('id順になっていない');
  if (!l[0].title || !l[0].version) throw new Error('カタログの中身が足りない');
});

/* ★2026-09-05★ 給与(kyuyo/)を Exally から 外した（給与は Rakunally）。
   ★実物の オペ（payroll.monthly）は 給与と 一緒に 出て行きました★
   ⇒★Exally に 実物の オペが 在るなら それで 試す／無ければ「未測定」と 言う★
   ⇒★★黙って 消さない★★＝「オペが 0本」を 緑に して しまうと、
     ★オペの 仕組みが 壊れても 誰も 気づかない★ */
T('★実物のオペが登録でき、契約の5つを返す（Exally に 在る 分）', () => {
  const 置き場 = path.join(ROOT, 'ops');
  const 在る = fs.existsSync(置き場)
    ? fs.readdirSync(置き場).filter((f) => /\.js$/.test(f)) : [];
  if (!在る.length) {
    console.log('  ★未測定★ Exally に 実物の オペが 在りません'
      + '（給与と 一緒に Rakunally へ 出ました）＝0本と 混ぜない');
    return;
  }
  R._reset();
  const real = require_(path.join(置き場, 在る[0]));
  R.register(real);
  const res = R.get(real.id).engine({});             // わざと空＝検証NGの道
  for (const k of ['value', 'cells', 'warnings', 'errors', 'provenance']) if (!(k in res)) throw new Error(k + ' が無い');
  if (res.value !== null) throw new Error('検証NGなのに value を作っている');
  if (!res.provenance.op) throw new Error('provenance が無い');
});

R._reset();
console.log('\n' + pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
