/* run.js — Exally のテストを全部走らせる(依存ゼロ・node だけ)
 *   node tests/run.js
 * 各テストファイルは自分で実行して、失敗があれば exit 1 を返す約束。
 */
'use strict';
const { execFileSync } = require('child_process');
const path = require('path');

const FILES = [
  'stamp.test.mjs',         // キャッシュバスター(?v=)の道具そのもの
  'suite-data.test.js',     // E0 共有データ層の契約
  'aggregate.test.js',      // E1 事業別集計(純関数)
  'ledger-source.test.js',  // E2 台帳→期間の実績値(ctx)
  'cross-agg.test.js',      // E5 横断集計(事業別のまとめ)
  'hub-ui.mjs',             // E1 UI 全ボタン(jsdom)
  'grid-xlsx.test.mjs',     // ★グリッド→xlsx の変換と「落ちる物」の警告(+ toHFVal との同期)
  'excel-version.test.mjs', // ★その式が相手のExcelで動くか(Excelに無い23個＝常時 / 版マーカー14個＝版連動)
  ['excel-version.test.mjs', '--self-test'],
  // ★staging(GitHub Pages のサブパス配信)で壊れる書き方＋本番倉庫への誤接続を止める恒久ガード
  'pages-hosting.test.mjs',
  ['pages-hosting.test.mjs', '--self-test'],  // ★わざと壊して赤になるかの自己確認(7通り)
  // P1② 版対応 検証ハーネス
  'xlsx-harness/roundtrip.test.mjs',        // 数式入りxlsxの往復(SheetJS・★新関数の _xlfn.)
  'xlsx-harness/bare-form.test.mjs',   // ★「客が最初に書く形」のケースが無い関数を赤にする(R19の再発防止)
  'xlsx-harness/alias.test.mjs',            // ★日本語UI名→本名(JIS→DBCS / YEN→DOLLAR。入口=エンジン/出口=書き出し)
  'xlsx-harness/xlfn-coverage.test.mjs',    // ★書き出す関数名が分類済みか(_xlfn.の付け忘れを止める)
  'xlsx-harness/version-scope.test.mjs',    // ★「版対応はここまで」の記述と実装がズレたら赤
  'xlsx-harness/compare.mjs',               // Excelの真値と突合(新規の不一致があれば赤)
  ['xlsx-harness/compare.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認
  ['xlsx-harness/nesting-audit.mjs', '--probe', '--check'] // ★入れ子で壊れる式が増えていないか
];

let ng = 0;
for (const f of FILES) {
  const [file, ...args] = Array.isArray(f) ? f : [f];
  console.log('\n=== ' + file + (args.length ? ' ' + args.join(' ') : '') + ' ===');
  try { execFileSync(process.execPath, [path.join(__dirname, file), ...args], { stdio: 'inherit' }); }
  catch (e) { ng++; }
}
console.log('\n' + (ng ? '★ ' + ng + ' ファイルで失敗' : '全テストファイル 緑'));
process.exit(ng ? 1 : 0);
