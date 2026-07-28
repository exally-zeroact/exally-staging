/* run.js — Exally のテストを全部走らせる(依存ゼロ・node だけ)
 *   node tests/run.js
 * 各テストファイルは自分で実行して、失敗があれば exit 1 を返す約束。
 */
'use strict';
const { execFileSync } = require('child_process');
const path = require('path');

const FILES = [
  'suite-data.test.js',     // E0 共有データ層の契約
  'aggregate.test.js',      // E1 事業別集計(純関数)
  'access-drift.test.js',   // E1 利用権 + Kyually版とのドリフト突合
  'periods-drift.test.js',  // E2 締め方(期間) + Kyually版との全パターン突合
  'ledger-agg.test.js',     // E2 台帳→期間の実績値(ctx)
  'hub-ui.mjs'              // E1 UI 全ボタン(jsdom)
];

let ng = 0;
for (const f of FILES) {
  console.log('\n=== ' + f + ' ===');
  try { execFileSync(process.execPath, [path.join(__dirname, f)], { stdio: 'inherit' }); }
  catch (e) { ng++; }
}
console.log('\n' + (ng ? '★ ' + ng + ' ファイルで失敗' : '全テストファイル 緑'));
process.exit(ng ? 1 : 0);
