/* periods-drift.test.js — 期間定義(締め方)のテスト ＋ ★Kyually版との全パターン突合★
 *
 * なぜ必要か: 期間の定義は Kyually(lib/periods.js)が唯一の源。別リポジトリで import できないため
 *   Exally にも実体があるが、片方だけ直すと「台帳の期間」と「給料の期間」がズレて
 *   支払いが1期間ぶん抜ける/二重になる。＝最賃38県のコピペ・ドリフト事故と同じ形。
 * そこで 4方式 × 月末28/29/30/31 × N × 全日付 で両者を突き合わせ、1つでもズレたら赤にする。
 * payslip-app が無い環境では ★スキップしたと明示★ する(黙って緑にしない)。
 */
'use strict';
const assert = require('assert');
const fs = require('fs');
const path = require('path');
const P = require('../lib/periods.js');

const T = [];
function test(name, fn) { T.push({ name, fn }); }
let skipped = 0;

/* ═══ 期間の割り方そのもの(実数値) ═══ */

test('10日締め(代行)=1〜10/11〜20/21〜末 の3期間', () => {
  const ps = P.buildPeriods('2026-07', 'ten');
  assert.strictEqual(ps.length, 3);
  assert.deepStrictEqual(ps.map(p => p.key), ['P1', 'P2', 'P3']);
  assert.deepStrictEqual(ps.map(p => [p.from, p.to]), [
    ['2026-07-01', '2026-07-10'], ['2026-07-11', '2026-07-20'], ['2026-07-21', '2026-07-31']
  ]);
  assert.deepStrictEqual(ps.map(p => p.label), ['1〜10', '11〜20', '21〜末']);
  assert.deepStrictEqual(ps.map(p => p.days), [10, 10, 11]);
});

test('半月=1〜15/16〜末', () => {
  const ps = P.buildPeriods('2026-07', 'half');
  assert.deepStrictEqual(ps.map(p => [p.from, p.to]), [['2026-07-01', '2026-07-15'], ['2026-07-16', '2026-07-31']]);
});

test('月まとめ=1期間だけ(キーは P1・M ではない)', () => {
  const ps = P.buildPeriods('2026-07', 'monthly');
  assert.strictEqual(ps.length, 1);
  assert.strictEqual(ps[0].key, 'P1', 'E0契約の period キーはこの値をそのまま使う');
  assert.strictEqual(ps[0].label, '1〜末');
  assert.deepStrictEqual([ps[0].from, ps[0].to], ['2026-07-01', '2026-07-31']);
});

test('月末が正しい(28/29/30/31・うるう年)', () => {
  assert.strictEqual(P.buildPeriods('2026-02', 'ten')[2].to, '2026-02-28');
  assert.strictEqual(P.buildPeriods('2024-02', 'ten')[2].to, '2024-02-29', 'うるう年');
  assert.strictEqual(P.buildPeriods('2026-04', 'ten')[2].to, '2026-04-30');
  assert.strictEqual(P.buildPeriods('2026-01', 'ten')[2].to, '2026-01-31');
});

test('任意N日=端数は最後の期間に入る(日が漏れない)', () => {
  const ps = P.buildPeriods('2026-07', 'ndays', 7);   // 31日 / 7日 = 4期間+3日
  assert.strictEqual(ps.length, 5);
  assert.strictEqual(ps[0].from, '2026-07-01');
  assert.strictEqual(ps[4].to, '2026-07-31');
  // ★全期間の日数の合計が その月の日数と一致する(1日も漏れない・重複しない)
  assert.strictEqual(ps.reduce((s, p) => s + p.days, 0), 31);
});

test('どの締め方でも 月の全日がちょうど1つの期間に入る(漏れ・重複ゼロ)', () => {
  const cases = [['2026-01', 'ten', 0], ['2026-02', 'half', 0], ['2024-02', 'ndays', 3],
                 ['2026-04', 'ndays', 30], ['2026-07', 'monthly', 0], ['2026-11', 'ndays', 1]];
  for (const [ym, method, n] of cases) {
    const last = P.lastDayOf(+ym.slice(0, 4), +ym.slice(5, 7));
    for (let d = 1; d <= last; d++) {
      const ymd = ym + '-' + ('0' + d).slice(-2);
      const hit = P.buildPeriods(ym, method, n).filter(p => ymd >= p.from && ymd <= p.to);
      assert.strictEqual(hit.length, 1, ymd + ' (' + method + n + ') が ' + hit.length + ' 個の期間に入った');
      assert.strictEqual(P.periodKeyOf(ymd, ym, method, n), hit[0].key);
    }
  }
});

test('期間外の日付は null(勝手にどこかへ入れない)', () => {
  assert.strictEqual(P.periodKeyOf('2026-08-01', '2026-07', 'ten'), null);
  assert.strictEqual(P.periodKeyOf('2026-06-30', '2026-07', 'ten'), null);
});

test('知らない締め方は monthly に倒す(落ちない)', () => {
  assert.strictEqual(P.buildPeriods('2026-07', 'unknown').length, 1);
  assert.strictEqual(P.buildPeriods('2026-07', undefined).length, 1);
  assert.strictEqual(P.buildPeriods('2026-07', null).length, 1);
});

test('壊れた ym は空配列(例外にしない)', () => {
  assert.deepStrictEqual(P.buildPeriods('', 'ten'), []);
  assert.deepStrictEqual(P.buildPeriods('2026-13', 'ten'), []);
  assert.deepStrictEqual(P.buildPeriods('abcd-ef', 'ten'), []);
});

test('hasSplit: 分割ありの判定', () => {
  assert.strictEqual(P.hasSplit('monthly'), false);
  ['half', 'ten', 'ndays'].forEach(m => assert.strictEqual(P.hasSplit(m), true, m));
});

/* ═══ ★Kyually版との全パターン突合★ ═══ */

const KY = path.join(__dirname, '..', '..', 'payslip-app', 'lib', 'periods.js');

test('★ドリフト突合: Kyually版と全パターンで一致する', () => {
  if (!fs.existsSync(KY)) {
    skipped++;
    console.log('       (このマシンに payslip-app が無いためスキップ: ' + KY + ')');
    return;
  }
  const K = require(KY);
  const YMS = ['2024-02', '2026-01', '2026-02', '2026-04', '2026-06', '2026-07', '2026-11', '2026-12', '2100-02'];
  const NS = [1, 2, 3, 5, 7, 10, 13, 15, 20, 30, 31, 45];
  let compared = 0;
  const diffs = [];

  for (const ym of YMS) {
    for (const method of ['monthly', 'half', 'ten', 'ndays']) {
      const ns = method === 'ndays' ? NS : [0];
      for (const n of ns) {
        const mine = P.buildPeriods(ym, method, n);
        const theirs = K.buildPeriods(ym, method, n);
        compared++;
        if (JSON.stringify(mine) !== JSON.stringify(theirs)) {
          diffs.push(ym + '/' + method + '/' + n + '\n         Exally =' + JSON.stringify(mine) + '\n         Kyually=' + JSON.stringify(theirs));
          continue;
        }
        // その月の全日で periodKeyOf も突合
        const last = P.lastDayOf(+ym.slice(0, 4), +ym.slice(5, 7));
        for (let d = 1; d <= last; d++) {
          const ymd = ym + '-' + ('0' + d).slice(-2);
          compared++;
          if (P.periodKeyOf(ymd, ym, method, n) !== K.periodKeyOf(ymd, ym, method, n)) {
            diffs.push(ymd + '/' + method + '/' + n + ' → Exally=' + P.periodKeyOf(ymd, ym, method, n) + ' / Kyually=' + K.periodKeyOf(ymd, ym, method, n));
          }
        }
      }
    }
  }
  // METHODS 一覧そのものも一致(片方に締め方を足したら赤)
  assert.deepStrictEqual(P.METHODS.slice().sort(), K.METHODS.slice().sort(), 'METHODS がズレている');
  assert.strictEqual(diffs.length, 0, diffs.length + '件ドリフト:\n       ' + diffs.slice(0, 5).join('\n       '));
  assert.ok(compared > 3000, '突合した数が少なすぎる: ' + compared);
  console.log('       (' + compared.toLocaleString('ja-JP') + ' パターンを突合して全一致)');
});

test('★ファイル自体が Kyually 版と同一である(コメント以外の差が無い)', () => {
  if (!fs.existsSync(KY)) { skipped++; console.log('       (payslip-app が無いためスキップ)'); return; }
  const strip = (s) => s.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*$/gm, '').replace(/\s+/g, ' ').trim();
  const mine = strip(fs.readFileSync(path.join(__dirname, '..', 'lib', 'periods.js'), 'utf8'));
  const theirs = strip(fs.readFileSync(KY, 'utf8'));
  assert.strictEqual(mine, theirs, 'コードがズレている(Kyually側を直してから複製し直すこと)');
});

/* ═══ 実行 ═══ */
(async () => {
  let ng = 0;
  for (const t of T) {
    try { await t.fn(); console.log('  ok   ' + t.name); }
    catch (e) { ng++; console.log('  NG   ' + t.name + '\n       ' + (e && e.message)); }
  }
  console.log('\nperiods: ' + (T.length - ng) + '/' + T.length + ' passed' + (skipped ? '  ★' + skipped + '件はスキップ(緑ではない)' : ''));
  if (ng) process.exit(1);
})();
