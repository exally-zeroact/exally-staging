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
  'typed-value.test.mjs',    // ★E3: 1,234 が文字列で合計に入らない(金が落ちる・期限9/30)
  ['typed-value.test.mjs', '--self-test'],
  'grid-refedit.test.mjs',   // ★書き間違えた式を直せるか(=B1+30 の B1 を A1 に直せる／数字の直後で式を壊さない)
  ['grid-refedit.test.mjs', '--self-test'],
  'grid-edit-ui.mjs',        // ★本物の book.html を読み込んで、本物の insertRefAddr を動かす(画面の中で直せるか)
  'excel-shortcuts.test.mjs',             // ★Excelと同じキー割り当てを 本物の画面に実際に押して確かめる(真値は実Excelから機械で取った)
  ['excel-shortcuts.test.mjs', '--self-test'],
  'mobile-labels.test.mjs',               // ★スマホの幅で 字を消して「絵だけ」にするのを禁じる(司さんのiPhoneで 📂💾📊 の絵だけになっていた)
  ['mobile-labels.test.mjs', '--self-test'],
  'grid-sort.test.mjs',                   // ★並べ替え(実Excelを COM で動かして測った並び順・見出し判定・式の運ばれ方)
  ['grid-sort.test.mjs', '--self-test'],
  'grid-filter.test.mjs',                 // ★絞り込み(実Excelで実測: 行は消えず隠れるだけ・★合計は変わらない★)
  ['grid-filter.test.mjs', '--self-test'],
  'grid-freeze.test.mjs',                 // ★ウィンドウ枠の固定(実Excelで実測)＋描き方を触った後に 前からある物が壊れていないか
  ['grid-freeze.test.mjs', '--self-test'],
  'grid-find.test.mjs',                   // ★検索と置換(実Excelで実測: ★置換は式を見る＝答えが変わる★／* ? ~ のワイルドカード)
  ['grid-find.test.mjs', '--self-test'],
  'grid-print.test.mjs',                  // ★印刷(実Excelの既定＝A4縦・余白・枠線なし)＋★白紙の印刷ダイアログを出さない★
  ['grid-print.test.mjs', '--self-test'],
  'grid-valid.test.mjs',                  // ★入力の決まり(一覧から選ぶ/整数の範囲)＝打った時だけ止める・合っていない値を数える
  ['grid-valid.test.mjs', '--self-test'],
  'grid-stats.test.mjs',                  // ★選んだ所の合計・平均・個数(帯)＝黙って小さい合計を出さない
  ['grid-stats.test.mjs', '--self-test'],
  'ctx-menu.test.mjs',                    // ★右クリックが画面の中に収まる(743pxが619pxの画面で 上へ470px はみ出した事故)
  ['ctx-menu.test.mjs', '--self-test'],
  'cond-format.test.mjs',                 // ★条件付き書式の当たり判定(実Excelの真値と突き合わせ)
  ['cond-format.test.mjs', '--self-test'],
  'cond-format-ui.test.mjs',              // ★本物の画面で 実際に押す(部品が緑=画面で使える ではない)
  ['cond-format-ui.test.mjs', '--self-test'],
  'login-gate.test.mjs',                  // ★表の画面にもログイン／忘れた人の逃げ道（無いと二度と入れない）
  ['login-gate.test.mjs', '--self-test'],
  'ai-reason.test.mjs',                   // ★AIに繋がらない時の理由と次の一手／★空のセルでAIを呼ばない(お金)★
  ['ai-reason.test.mjs', '--self-test'],
  'no-dead-ui.test.mjs',     // ★出来ていない物のボタン/画面を止める窓/中の言葉(STEP6・実装予定)を客に見せない
  ['no-dead-ui.test.mjs', '--self-test'],
  'word-export-import.test.mjs',   // ★言い方を「書き出す↔読み込む」に固定
  ['word-export-import.test.mjs', '--self-test'],
  'excel-parity.test.mjs',         // ★Excelとの差を機械で数え直す(表が古くなったら赤)
  ['excel-parity.test.mjs', '--self-test'],
  'cross-sheet.test.mjs',          // ★他のシートを参照している合計が黙って小さくならないか(527,000が186,000)
  'smart-rounding.test.mjs',       // ★計算の結果を14桁で丸めさせない(消費税が1円ズレる)
  ['smart-rounding.test.mjs', '--self-test'],
  'text-format.test.mjs',          // ★TEXT()の書式コード(曜日aaa)。実物730本がシリアル値のまま出ていた
  ['text-format.test.mjs', '--self-test'],
  'no-silent-optional.test.mjs',   // ★typeofで守って「無ければ黙って素通り」を許さない
  ['no-silent-optional.test.mjs', '--self-test'],
  'book-open.test.mjs',      // ★受け取ったブックを「開いて何も変えずに保存」しても1バイトも変わらない(zip直編集の3本＋book-open.js)
  'diff-preview.test.mjs',   // ★直す前に必ず見せる(方針ver.6の②)。1直しで3シート18本 書き換わる
  ['diff-preview.test.mjs', '--self-test'],
  'table-refs.test.mjs',     // ★表の名前での参照(Table[列名])→A1範囲。実物の式11,669本が1本残らず#ERRORだった
  ['table-refs.test.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認(16通り)
  'no-duplicate-libs.test.mjs', // ★同じ物を2箇所に置かせない(法定データのコピペ・ドリフト防止)
  'excel-version.test.mjs', // ★その式が相手のExcelで動くか(Excelに無い23個＝常時 / 版マーカー14個＝版連動)
  ['excel-version.test.mjs', '--self-test'],
  // ★staging(GitHub Pages のサブパス配信)で壊れる書き方＋本番倉庫への誤接続を止める恒久ガード
  'env-badge.test.mjs',      // ★テスト環境の帯(本番に出さない・全画面に入っている)
  ['env-badge.test.mjs', '--self-test'],
  'pages-hosting.test.mjs',
  ['pages-hosting.test.mjs', '--self-test'],  // ★わざと壊して赤になるかの自己確認(7通り)
  'refs-resolve.test.mjs',      // ★読んでいるファイルが実在するか(require/importも参照として数える)
  ['refs-resolve.test.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認
  'api-claude.test.mjs',        // ★チャットが客に言う基準数値(実数リテラル・NaN混入検知)
  ['api-claude.test.mjs', '--self-test'],  // ★失敗しても200で「答えのふり」をしていた穴(2026-08-22)
  'no-hardcoded-statutory.test.mjs',      // ★法定の率・額を配信物の文に直書きさせない(説明文だけ年度で取り残される事故)
  ['no-hardcoded-statutory.test.mjs', '--self-test'], // ★わざと壊して赤になるか＋誤検知が出ないか
  'no-hardcoded-supa.test.mjs',           // ★倉庫の向き先を js/supa-config.js 以外に書かせない(テストrepoが本番倉庫を触る事故)
  ['no-hardcoded-supa.test.mjs', '--self-test'], // ★わざと壊して赤になるか＋誤検知が出ないか
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
const 失敗一覧 = [];
for (const f of FILES) {
  const [file, ...args] = Array.isArray(f) ? f : [f];
  console.log('\n=== ' + file + (args.length ? ' ' + args.join(' ') : '') + ' ===');
  try { execFileSync(process.execPath, [path.join(__dirname, file), ...args], { stdio: 'inherit' }); }
  catch (e) {
    ng++;
    /* ★落ちた理由を必ず出す（2026-08-23）★ CIで1回 赤→同じコミットを回し直したら緑＝★ムラ★だった。
       「N ファイルで失敗」だけでは ★中で殺されたのか／自分で1を返したのか★ が分からない。 */
    const 印 = e && e.signal ? ('★中で殺された(signal=' + e.signal + ')★＝新しい壊れではない可能性')
      : ('自分で ' + (e && e.status !== undefined && e.status !== null ? e.status : '?') + ' を返した');
    console.log('  ★落ちた★ ' + file + (args.length ? ' ' + args.join(' ') : '') + ' … ' + 印);
    失敗一覧.push(file + (args.length ? ' ' + args.join(' ') : '') + '（' + 印 + '）');
  }
}
console.log('\n' + (ng ? '★ ' + ng + ' ファイルで失敗' : '全テストファイル 緑'));
for (const 名 of 失敗一覧) console.log('   ・' + 名);
process.exit(ng ? 1 : 0);
