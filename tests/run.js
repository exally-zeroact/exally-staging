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
  'grid-xlsx.test.mjs',  // ★グリッド→xlsx の変換と「落ちる物」の警告(+ toHFVal との同期)
  'grid-date.test.mjs',    // ★打った日付が日付として計算できるか(+30が2056にならない)＋数を日付に化けさせない
  ['grid-date.test.mjs', '--self-test'],
  'grid-colwidth.test.mjs',  // ★渡した相手の画面で ######## にならないか(日付の列に幅を付ける)
  ['grid-colwidth.test.mjs', '--self-test'],
  'grid-refedit.test.mjs',   // ★書き間違えた式を直せるか(=B1+30 の B1 を A1 に直せる／数字の直後で式を壊さない)
  ['grid-refedit.test.mjs', '--self-test'],
  'no-duplicate-libs.test.mjs', // ★同じ物を2箇所に置かせない(法定データのコピペ・ドリフト防止)
  'excel-version.test.mjs', // ★その式が相手のExcelで動くか(Excelに無い23個＝常時 / 版マーカー14個＝版連動)
  ['excel-version.test.mjs', '--self-test'],
  // ★staging(GitHub Pages のサブパス配信)で壊れる書き方＋本番倉庫への誤接続を止める恒久ガード
  'pages-hosting.test.mjs',
  ['pages-hosting.test.mjs', '--self-test'],  // ★わざと壊して赤になるかの自己確認(7通り)
  'refs-resolve.test.mjs',      // ★読んでいるファイルが実在するか(require/importも参照として数える)
  ['refs-resolve.test.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認
  'api-claude.test.mjs',        // ★チャットが客に言う基準数値(実数リテラル・NaN混入検知)
  'no-hardcoded-statutory.test.mjs',      // ★法定の率・額を配信物の文に直書きさせない(説明文だけ年度で取り残される事故)
  ['no-hardcoded-statutory.test.mjs', '--self-test'], // ★わざと壊して赤になるか＋誤検知が出ないか
  'ios-unsupported.test.mjs',   // ★iPhoneで動かない書き方(type=month/octet-stream/writeFile/Blob散在)
  ['ios-unsupported.test.mjs', '--self-test'],
  'op-registry.test.mjs',       // ★契約の入口(二重登録は投げる)
  'op-boundary.test.mjs',       // ★契約の線(⑤呼ばれているか/⑧面を呼び返していないか/provenance必須)
  ['op-boundary.test.mjs', '--self-test'],
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
