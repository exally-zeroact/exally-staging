/* access-drift.test.js — 利用権の判定テスト ＋ ★Kyually版とのドリフト突合★
 *
 * なぜ必要か: 同じ「使える/停止」の判定が Exally と Kyually の2箇所に実体を持つ(別リポジトリで
 *   import できないため)。片方だけ直すと「停止したはずのアカウントが使える」事故になる。
 *   ＝過去に最賃38県が誤値になったコピペ・ドリフト事故と同じ形。
 * そこで全plan値で両者の判定を突き合わせ、1つでもズレたら赤にする。
 * payslip-app が同じマシンに無い環境では ★スキップしたと明示★ する(黙って緑にしない)。
 */
'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const Access = require('../lib/access.js');

const T = [];
function test(name, fn) { T.push({ name, fn }); }
let skipped = 0;

/* ═══ Exally側の判定そのもの ═══ */

test('行が無い(初回)は使える', () => {
  assert.deepStrictEqual(Access.accessState(null), { ok: true, reason: 'new' });
  assert.deepStrictEqual(Access.accessState(undefined), { ok: true, reason: 'new' });
});

test('trial / paid / free は使える', () => {
  assert.deepStrictEqual(Access.accessState({ plan: 'trial' }), { ok: true, reason: 'trial' });
  assert.deepStrictEqual(Access.accessState({ plan: 'paid' }), { ok: true, reason: 'paid' });
  assert.deepStrictEqual(Access.accessState({ plan: 'free' }), { ok: true, reason: 'free' });
});

test('disabled は止める', () => {
  assert.deepStrictEqual(Access.accessState({ plan: 'disabled' }), { ok: false, reason: 'disabled' });
});

test('知らない plan は「止める」側に倒す(安全側)', () => {
  assert.deepStrictEqual(Access.accessState({ plan: 'unknown' }), { ok: false, reason: 'disabled' });
  assert.deepStrictEqual(Access.accessState({ plan: 'PAID' }), { ok: false, reason: 'disabled' }, '大文字は別物として止める');
  assert.deepStrictEqual(Access.accessState({ plan: ' trial' }), { ok: false, reason: 'disabled' }, '前後の空白も別物');
});

test('plan が空なら trial 扱い(行はあるが未設定)', () => {
  assert.deepStrictEqual(Access.accessState({}), { ok: true, reason: 'trial' });
  assert.deepStrictEqual(Access.accessState({ plan: '' }), { ok: true, reason: 'trial' });
  assert.deepStrictEqual(Access.accessState({ plan: null }), { ok: true, reason: 'trial' });
});

test('停止画面の文言は冷たくない・空でない', () => {
  const m = Access.lockMessage();
  assert.ok(m && typeof m.title === 'string' && m.title.length > 0);
  assert.strictEqual(/管理者に連絡|お問い合わせください/.test(m.title + m.body), false, '冷たい文言は使わない');
});

/* ═══ ★Kyually版とのドリフト突合★ ═══ */

const KY = path.join(__dirname, '..', '..', 'payslip-app', 'lib', 'access.js');

test('★ドリフト突合: Kyually版と全plan値で判定が一致する', () => {
  if (!fs.existsSync(KY)) {
    skipped++;
    console.log('       (このマシンに payslip-app が無いためスキップ: ' + KY + ')');
    return;
  }
  const KAccess = require(KY);
  // 実在する値・境界・壊れた値を網羅
  const CASES = [
    null, undefined, {}, { plan: '' }, { plan: null }, { plan: undefined },
    { plan: 'trial' }, { plan: 'paid' }, { plan: 'free' }, { plan: 'disabled' },
    { plan: 'unknown' }, { plan: 'PAID' }, { plan: ' trial' }, { plan: 'trial ' },
    { plan: 'Disabled' }, { plan: 0 }, { plan: false }, { plan: 'expired' },
    { plan: 'trial', expires_at: '2020-01-01T00:00:00Z' },   // 期限切れ列は未参照のはず(両者とも)
    { plan: 'paid', expires_at: null }
  ];
  const diffs = [];
  CASES.forEach(function (c) {
    const mine = Access.accessState(c);
    const theirs = KAccess.accessState(c);
    if (JSON.stringify(mine) !== JSON.stringify(theirs)) {
      diffs.push(JSON.stringify(c) + ' → Exally=' + JSON.stringify(mine) + ' / Kyually=' + JSON.stringify(theirs));
    }
  });
  assert.strictEqual(diffs.length, 0, 'ドリフトしています:\n       ' + diffs.join('\n       '));
  // 使える plan の一覧そのものも一致していること(片方に plan を足したら赤にする)
  assert.deepStrictEqual(Access.PLANS.slice().sort(), KAccess.PLANS.slice().sort(), 'PLANS 一覧がズレている');
});

/* ═══ 実行 ═══ */
(async () => {
  let ng = 0;
  for (const t of T) {
    try { await t.fn(); console.log('  ok   ' + t.name); }
    catch (e) { ng++; console.log('  NG   ' + t.name + '\n       ' + (e && e.message)); }
  }
  console.log('\naccess: ' + (T.length - ng) + '/' + T.length + ' passed' + (skipped ? '  ★' + skipped + '件はスキップ(緑ではない)' : ''));
  if (ng) process.exit(1);
})();
