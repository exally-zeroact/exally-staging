/* run.js — Exally のテストを全部走らせる(依存ゼロ・node だけ)
 *   node tests/run.js
 * 各テストファイルは自分で実行して、失敗があれば exit 1 を返す約束。
 */
'use strict';
const { execFileSync } = require('child_process');
const path = require('path');

const FILES = [
  'stamp.test.mjs',         // キャッシュバスター(?v=)の道具そのもの
  ['sql-guard.test.mjs', '--self-test'],  // ★倉庫にSQLを当てる門番（本番の実データが同居している）★
  ['empty-ref-zero.test.mjs', '--self-test'],  // ★式が空セルを指したら Excelは0（実物2,918本が これで合う）★
  ['shini-code.test.mjs'],                     // ★呼ばれない 物を 置いたままに しない（2026-09-11 下請け4本を 消した）★
  ['shiki-kiru.test.mjs'],
  ['shiki-katachi.test.mjs'],
  ['shiki-sansho.test.mjs'],                   // ★参照を 読む（地図を 作る 前に 要る）★                  // ★切った かたまりを 形に する（土台を 自作する 2枚目）★                     // ★式を 字の かたまりに 切る（土台を 自作する 1枚目）★
  ['shiki-chizu.test.mjs'],                    // ★頼りの 地図＝どの 順に 計算するか★
  ['shiki-keisan.test.mjs'],                   // ★つなぎ（＋−×÷＾＆ 大小 ％）＝実Excel 162通り★
  ['shiki-kansuu.test.mjs'],                   // ★四角（A1:A3）と 関数＝実Excel 37通り（土台④の 後半）★
  ['shiki-hyou.test.mjs'],                     // ★本体＝表を持ち 依存を追い 変わった所だけ直す（借り物の 55.2%）★
  ['shiki-sashikomi.test.mjs'],                // ★行/列の 入れ消しで 式が 追従する＝実Excel 17通り★
  ['shiki-afure.test.mjs'],                    // ★溢れ＝1つの式が 何マスにも 広がる（実Excel 16通り）★
  ['shiki-tsunagi.test.mjs'],                  // ★自前の 土台に 自前の 61個を 繋ぐ 皮（借り物を 1行も 通さない）★
  ['shiki-tsunagi.test.mjs', '--self-test'],   // ★皮を 外したら 赤に なるか／形を 潰したら 潰れるか★
  ['shiki-hyou-afure.test.mjs'],               // ★本体が 溢れ（こぼれ）を 置ける＝実Excel 16通り（土台⑤）★
  ['shiki-hyou-afure.test.mjs', '--self-test'],// ★材料を 変えたら 答えも 変わるか／紙に 無い 式を 押して いないか★
  ['shiki-hyou-omosa.test.mjs'],               // ★表が 大きく なっても N の 2乗に ならない（2026-09-14 に 2乗だった）★
  ['shiki-hyou-sakuin.test.mjs'],  // ★在る マスの 索引が いつも 合って いるか★（空の マスを 5,700万回 訊いて いた・2026-09-26）
  ['shiki-hyou-matomete.test.mjs'],  // ★まとめて 打っても 1つずつ と 答えが 同じか★（1打ち目に 式を 58万回 計算して いた・2026-10-05）
  ['shiki-hyou-matomete.test.mjs', '--self-test'],  // ★終わりを 呼び忘れると 古い 答えで 赤に なるか★
  ['afure-no-gamen-webkit.mjs'],  // ★溢れる 式の 画面の 字が 実Excel と 同じか★（物差し 137・57行・既知 3・2026-10-05）
  ['afure-no-gamen-webkit.mjs', '--self-test'],  // ★既知の 名簿を 空に すると 赤（道具が 空振りして いない）★
  ['hyou-no-kihon-no-ji-webkit.mjs'],  // ★General の 桁は 本の 既定の 字で 決める★（そのマスの 字で 決めて いた・2026-09-26）
  ['shiki-hyou-omosa.test.mjs', '--self-test'],// ★測りが 空振りして いないか（0本・0ミリ秒で 緑に しない）★
  ['shiki-hyou-hidzuke.test.mjs'],             // ★字の 日付・時刻を 数に する（実測 "12:30"→0.5208･･･）★
  ['shiki-hyou-hidzuke.test.mjs', '--self-test'],// ★1900年は 未測定と 書いて 在るか／時刻の 割合★
  ['shiki-hyou-hishigata.test.mjs'],           // ★菱形の頼りで古い答えが残らない（実物 37/176本→0/0）★
  ['shiki-hyou-hishigata.test.mjs','--self-test'],// ★輪を壊していない／棚に直す前の数が在る★
  ['kami-no-zairyou.test.mjs'],                // ★紙が 古いまま 置かれるのを 止める（材料の指紋）★
  ['kami-no-zairyou.test.mjs','--self-test'],  // ★1バイト変わったら気づく／合う指紋は通す／消えた材料★
  ['shiki-hyou-itamatagi.test.mjs'],           // ★板またぎ＝菱形・輪・無い板・裸のA1はその式の板★
  ['shiki-hyou-itamatagi.test.mjs','--self-test'],// ★下見の数が紙に在る★
  ['shiki-hyou-kata-no-kuchi.test.mjs'],       // ★型つきの口＝置いた事が 見ている式に 伝わる／直書き0個★
  ['shiki-hyou-kata-no-kuchi.test.mjs','--self-test'],// ★変数に受けて逃げる形も捕まえる／覚書きは数えない★
  ['dkei6-honban.test.mjs'],                   // ★棚㉝の6か所＝実Excel 26件（多列かつ／多行または／記号／field超過／0件）★
  ['dkei6-honban.test.mjs','--self-test'],     // ★「まだ」は名指し＋訳つき／合ったら許しを外させる★
  ['dkei-honban.test.mjs'],                    // ★D系12個を 本番の道で 実Excel 21行と 突き合わせ（見出しが 紐付かない 形）★
  ['dkei-honban.test.mjs','--self-test'],      // ★穴の 場所と 直した 印が 字で 在る／見ていない 6か所は 棚★
  ['shiki-kansuu-kami.test.mjs'],              // ★台が知る87個を 手元の紙60枚で 押す（632行）＝書くたび 自動で増える★
  ['shiki-kansuu-kami.test.mjs','--self-test'],// ★押した行が0なら赤／「まだ」は名指し16件・訳つき★
  ['shiki-kansuu-dkei.test.mjs'],              // ★D系12個を台へ＝実Excel 21行＋17行（0件の決めが 関数ごとに 違う）★
  ['shiki-kansuu-dkei.test.mjs','--self-test'],// ★台が知るD系を 呼んで数える（12/12）★
  ['shiki-kansuu-dsum.test.mjs'],              // ★417個の うち 移した 最初の 1個＝DSUM（実Excel 14／2）★
  ['shiki-kansuu-dsum.test.mjs','--self-test'],// ★DSUMを 外したら 5本とも 赤に なるか★（★土台が 在る 今は 素通り＝記録★）
  ['shiki-sagasu.test.mjs'],                   // ★INDEX と MATCH＝実物の のべ 68%（実Excel 193本と 突き合わせ）★
  ['shiki-sagasu.test.mjs', '--self-test'],    // ★見つからない時／転け方 2種／打つマスで 変わる／型1は 二分探索★
  ['shoshiki.test.mjs'],                       // ★書式の台＝TEXT も 画面も 同じ 台（実Excel 192行と 一致）★
  ['shoshiki.test.mjs', '--self-test'],        // ★時の隣の m は 分／-0 を 出さない／1900年の起点 2つ／曜日は 割り算★
  ['xlsb-jitai.test.mjs'],                     // ★.xlsb の マスごとの 字体を 自分で 読む（借り物は くれない）★
  ['ketsugou-yomu.test.mjs'],                  // ★Excel の 結合した マスを 読む（前は 1組も 読んで いなかった）★
  ['migi-ni-ugokasu.test.mjs'],                // ★右へ動かすと画面が真っ白（幅の広い列が在ると当たりが行き過ぎる）★
  ['kumikomi-shoshiki.test.mjs'],              // ★国で 変わる 書式番号38／15桁に 丸めてから 出す（実物で 見つけた）★
  ['moji-no-kotae.test.mjs'],                  // ★文字を 返す 式が 数に 化けない（2026-09-11 実Excel 14本 一致）★
  ['hashira-haba.test.mjs'],                   // ★列の 幅を 実Excel と 同じに する（2026-09-11 実測 8列 差0.00）★
  ['filter-shape.test.mjs', '--self-test'],    // ★FILTERの形／_xlws.の印／範囲の鎖（実物52+10本が これで合う）★
  ['no-lookbehind.test.mjs', '--self-test'],  // ★後読み正規表現＝旧iOS Safariで かたまりが丸ごと動かない★
  ['ribbon.test.mjs', '--self-test'],        // ★リボン＝Excelと同じ配置／押して 届くかまで 見る★
  ['ribbon-label.test.mjs', '--self-test'], // ★札が 箱に 収まり 同じに 見えない（08-30 監査役の 差し戻し）★
  ['ribbon-scope.test.mjs', '--self-test'], // ★空の箱を (a)これから/(b)付けません/(c)出さない に 分ける★
  ['view-extras.test.mjs', '--self-test'],  // ★表示｜表示＝測り直しで 見つけた 組★
  ['type-to-edit.test.mjs', '--self-test'],  // ★セルを 選んで いきなり 打つ（08-30 いうえおあ）★
  ['unused-param.test.mjs', '--self-test'], // ★持っているのに 渡していない 口が 増えていないか★
  ['ribbon-keytips.test.mjs', '--self-test'], // ★Alt の キーが 実 Excel と 同じ順で 効くか★
  ['ctx-menu-items.test.mjs', '--self-test'], // ★右クリックの 中身が 実 Excel に 追いついているか★
  ['header-footer.test.mjs', '--self-test'], // ★紙の 上と 下に 入れる 字（印は 実 Excel に 刷らせて 測った）★
  ['ribbon-launcher.test.mjs', '--self-test'], // ★組の 右下の ↘（開く先の 窓は 既に 在る）★
  ['ribbon-context.test.mjs', '--self-test'],
  'bunbo-context.test.mjs',                 // ★コンテキストタブの分母が3か所で揃うか（09-09 文書に 188/265/235 が同居していた）
  ['bunbo-context.test.mjs','--self-test'], // ＝古い数は消さず「今の正本は235」の印を必ず付ける // ★コンテキストタブ 8タブ／235部品の 正本★
  ['formula-extra.test.mjs', '--self-test'], // ★足りない 関数 12個（答えは 実 Excel の 実測）★
  ['valign.test.mjs', '--self-test'],        // ★上下揃え＝実Excelの既定は中央（実測）★
  ['clear.test.mjs', '--self-test'],         // ★クリア3通り＝中身だけ/書式だけ/すべて（実測）★
  ['ribbon-features.test.mjs', '--self-test'],  // ★数式の表示／再計算／ズーム／重複の削除★
  ['named-ranges.test.mjs', '--self-test'],  // ★名前の定義＝日本語の名前をエンジンに渡す前に開く★
  ['split-columns.test.mjs', '--self-test'], // ★区切り位置＝実Excelの TextToColumns どおり★
  ['chart.test.mjs', '--self-test'],         // ★グラフ＝実Excelの既定（360x216・凡例は2本以上）★
  ['comment.test.mjs', '--self-test'],       // ★コメント＝実Excelどおり（作者つき・赤い印だけ）★
  ['page-setup.test.mjs', '--self-test'],    // ★ページ設定＝実Excelの既定（A4縦・枠線も見出しも刷らない）★
  ['protect.test.mjs', '--self-test'],       // ★シートの保護＝実Excelどおり（既定は全セルロック・保護して初めて効く）★
  ['decimal-painter.test.mjs', '--self-test'], // ★小数の増減（上限30桁は実測）／書式のコピー（中身は写さない）★
  ['shift-cells.test.mjs', '--self-test'],   // ★セルの挿入/削除＝実Excelどおり（選んだ列だけ動く）★
  ['cell-styles.test.mjs', '--self-test'],   // ★セルのスタイル＝実Excelから読んだ真値（BGR→RGB）★
  ['sort2.test.mjs', '--self-test'],         // ★並べ替え2つの鍵＝実Excelの並びと一致★
  ['view-toggle.test.mjs', '--self-test'],
  ['link.test.mjs', '--self-test'],
  ['select-all-statusbar.test.mjs', '--self-test'], // ★Ctrl+A（すべて選ぶ）で 帯が落ちない（170億マスを 積もうとして 死んでいた）★         // ★リンク＝実Excelどおり（#467886の下線）／javascript:とdata:は弾く★   // ★表示の切り替え＝実Excelの既定（見出しも枠線も出す）★
  ['chart-types.test.mjs', '--self-test'],   /* ★グラフの種類（凡例の出る/出ない・軸の数・四分位は実測）★ */
  ['sparkline.test.mjs', '--self-test'],     /* ★スパークライン＝実測（飾りなし・勝敗だけマイナス別・縦軸は1本ごと・空は隙間）★ */
  ['formula-bar-type.test.mjs', '--self-test'], /* ★数を入れると 数式バーが落ちて セル選択そのものが 出来なくなっていた★ */
  ['table.test.mjs', '--self-test'],         /* ★テーブル(Ctrl+T)＝実測／行列を足しても 付箋・リンク・表が 置いていかれない★ */
  ['symbol-form.test.mjs', '--self-test'],   /* ★記号と特殊文字（並びは未測定＝理由つき）／フォーム（実測＝まわりの塊）★ */
  ['formula-tab.test.mjs', '--self-test'],   /* ★数式タブ＝計算方法(実測 既定は自動)／選択範囲から名前(見出しは入らない)／無い関数を出さない★ */
  ['view-review-tab.test.mjs', '--self-test'], /* ★ズーム(実測10〜400)／ブックの保護(守るとシートが足せない)／許可範囲／ブックの数★ */
  ['data-tab.test.mjs', '--self-test'],      /* ★フラッシュフィル(1つの見本で覚える)／再適用／ゴールシーク(実測100回・0.001)★ */
  ['home-tab.test.mjs', '--self-test'],      /* ★ホーム＝フォント(在るか測る)／文字の向き(実測の数)／クリップボード24個★ */
  ['pivot.test.mjs', '--self-test'],         /* ★ピボット＝実測の形と 数（列なしは2列だけ）／読みで並べない事も 書く★ */
  ['data-tab2.test.mjs', '--self-test'],     /* ★統合／アウトライン(実測 段2・1段で隠れる)／スライサー(見た目は未測定)★ */
  ['forecast.test.mjs', '--self-test'],      /* ★予測シート＝実測(傾き20・切片80・7つ目220)と 同じ／点検／メモ★ */
  ['automation.test.mjs', '--self-test'],    /* ★自動化＝Officeスクリプトの見本と同じ中身を うちの道具で（10項目）★ */
  ['csv-in.test.mjs', '--self-test'],        /* ★CSV読み込み＝実測どおり（"･･･"の中のカンマ・""・中の改行）／化けたらShift_JIS★ */
  ['objects.test.mjs', '--self-test'],       /* ★浮かぶ物（画像・図形・テキスト）＝形の番号と重なりは実測・色は写さない★ */
  ['ink.test.mjs', '--self-test'],           /* ★描画（手書き）＝ペンの色と太さは 正本の 項目名どおり／線もシートの座標★ */
  ['arrange.test.mjs', '--self-test'],
  ['view-window.test.mjs', '--self-test'],
  ['no-missing-call.test.mjs', '--self-test'],
  'name-vs-body.test.mjs',                   /* ★「わざと壊して赤になるか」と名乗るのに 何も壊していない試験を 探す（08-30 監査役）★ */
  ['theme.test.mjs', '--self-test'],
  ['boolean.test.mjs', '--self-test'],
  ['insert-diagram.test.mjs', '--self-test'],
  ['data-conn.test.mjs', '--self-test'],
  ['review2.test.mjs', '--self-test'],
  ['xml-dev.test.mjs', '--self-test'],
  ['shape-map-time.test.mjs', '--self-test'],
  ['model-web-data.test.mjs', '--self-test'], /* ★3Dは対角線で合わせて どの向きでも はみ出さない／つなぐ相手は1つずつ★ */ /* ★手書き→図形12/12／升目の日本地図47／日付の通し番号は実測と7/7★ */       /* ★XML＝結ぶ前は読めない・出せない（実測）／2つの読み取りで答えを合わせる★ */       /* ★校閲＝実Excelが測った6語と同じ答え／日本語は通す／見やすさは4.5:1★ */     /* ★詳細設定＝実測4通りの行数／つなぎ台帳＝出来ない更新は そう言う★ */ /* ★図解＝型159・節5は実測／字は形の後に出す（重なる丸で隠れた）／アイコンと数式はうちで描く★ */       /* ★TRUE/FALSEは大文字で出す／SUMは足さない・＋は数になる（実Excel実測）★ */         /* ★テーマ＝12の役割・濃淡は実Excelと15/15一致（0〜240で切り捨て）／色は写さない★ */ /* ★呼んでいるのに 無い 働きを 機械で 探す（08-30 colName 事故）★ */ /* ★表示タブ＝表示の番号1/2/3・分割はPanes4・窓の名前は空白2つ／紙の切れ目は列8で実Excelと同じ★ */       /* ★配置・回転・まとめる＋拡大縮小印刷＝そろえた後の 数は 実測どおり／倍率は 紙に 本当に 効く★ */
  'suite-data.test.js',     // E0 共有データ層の契約
  'aggregate.test.js',      // E1 事業別集計(純関数)
  'ledger-source.test.js',  // E2 台帳→期間の実績値(ctx)
  'cross-agg.test.js',      // E5 横断集計(事業別のまとめ)
  'hub-ui.mjs',             // E1 UI 全ボタン(jsdom)
  'grid-xlsx.test.mjs',     // ★グリッド→xlsx の変換と「落ちる物」の警告(+ toHFVal との同期)
  'grid-date.test.mjs',     // ★打った日付が日付として計算できるか(+30が2056にならない)＋数を日付に化けさせない
  ['grid-date.test.mjs', '--self-test'],
  'grid-colwidth.test.mjs', // ★渡した相手の画面で ######## にならないか(日付の列に幅を付ける)
  ['grid-colwidth.test.mjs', '--self-test'],
  'typed-value.test.mjs',                 // ★E3: 1,234 が文字列で合計に入らない(金が落ちる・期限9/30)
  ['typed-value.test.mjs', '--self-test'],
  'no-duplicate-libs.test.mjs', // ★同じ物を2箇所に置かせない(法定データのコピペ・ドリフト防止)
  'refs-resolve.test.mjs',      // ★読んでいるファイルが実在するか(require/importも参照として数える)
  ['refs-resolve.test.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認
  'api-claude.test.mjs',        // ★チャットが客に言う基準数値(実数リテラル・NaN混入検知)
  ['api-claude.test.mjs', '--self-test'],  // ★失敗しても200で「答えのふり」をしていた穴(2026-08-22)
  'ios-unsupported.test.mjs',   // ★iPhoneで動かない書き方(type=month/octet-stream/writeFile/Blob散在)
  ['ios-unsupported.test.mjs', '--self-test'],
  'op-registry.test.mjs',       // ★契約の入口(二重登録は投げる)
  'op-boundary.test.mjs',       // ★契約の線(⑤呼ばれているか/⑧面を呼び返していないか/provenance必須)
  ['op-boundary.test.mjs', '--self-test'],
  'no-hardcoded-statutory.test.mjs',      // ★法定の率・額を配信物の文に直書きさせない(説明文だけ年度で取り残される事故)
  ['no-hardcoded-statutory.test.mjs', '--self-test'], // ★わざと壊して赤になるか＋誤検知が出ないか
  'no-hardcoded-supa.test.mjs',           // ★倉庫の向き先を js/supa-config.js 以外に書かせない(テストrepoが本番倉庫を触る事故)
  ['no-hardcoded-supa.test.mjs', '--self-test'], // ★わざと壊して赤になるか＋誤検知が出ないか
  'no-silent-optional.test.mjs',          // ★typeof で守って「無ければ黙って素通り」を許さない(14シート開いてもタブが1枚だった事故)
  ['no-silent-optional.test.mjs', '--self-test'],
  'smart-rounding.test.mjs',              // ★計算の結果を14桁で丸めさせない(消費税が1円ズレる)
  ['smart-rounding.test.mjs', '--self-test'],
  'text-format.test.mjs',                 // ★TEXT()の書式コード(曜日aaa/月名/時刻の分)。実物730本がシリアル値のまま出ていた
  ['text-format.test.mjs', '--self-test'],
  'no-dead-ui.test.mjs',                  // ★出来ていない物のボタン/画面を止める窓/中の言葉(STEP6・実装予定)を客に見せない
  ['no-dead-ui.test.mjs', '--self-test'],
  'cross-sheet.test.mjs',                 // ★他のシートを参照している合計が黙って小さくならないか(527,000が186,000になった)
  'table-refs.test.mjs',                  // ★表の名前での参照(Table[列名])→A1範囲。実物の式11,669本が1本残らず#ERRORだった
  ['table-refs.test.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認(16通り)
  'kobore.test.mjs',                      // ★足した関数が2つ以上返せるか＝本番に前から在った穴（エンジンに通して打つ）
  ['kobore.test.mjs', '--self-test'],     // ★大きさの宣言を外して 赤になるか
  'bahttext.test.mjs',                    // ★BAHTTEXT＝タイ語の金額。答えは実Excelに116通り打たせた真値
  ['bahttext.test.mjs', '--self-test'],   // ★本物の部品を11通り壊して 赤になるか
  'hyou-no-soto.test.mjs',                // ★診断2本目「ほかの表の その行を見ている」＝実物で62か所（直し方まで出す）
  ['hyou-no-soto.test.mjs', '--self-test'],// ★本物の部品を11通り壊して 赤になるか
  'book-open.test.mjs',                   // ★受け取ったブックを「開いて何も変えずに保存」しても壊れない(実機で出た事故)
  'diff-preview.test.mjs',                // ★直す前に必ず見せる(方針ver.6の②)。1直しで3シート18本 書き換わる
  ['diff-preview.test.mjs', '--self-test'],
  'word-export-import.test.mjs',          // ★言い方を「書き出す↔読み込む」に固定(同じ動きに2通りの名前を付けない)
  ['word-export-import.test.mjs', '--self-test'],
  'excel-parity.test.mjs',                // ★Excelとの差を機械で数え直す(表が古くなったら赤)
  ['excel-parity.test.mjs', '--self-test'],
  'grid-refedit.test.mjs',                // ★式の中の参照を直せるか(=B1+30 の B1 を A1 に直す)
  ['grid-refedit.test.mjs', '--self-test'],
  'grid-edit-ui.mjs',                     // ★本物の book.html で insertRefAddr を動かす(画面の中で直せるか)
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
  'ctx-menu-a-ui.test.mjs',               // ★本物の画面で 右クリックを出す(灰・一覧が長い時の 切り替え)
  ['ctx-menu-a-ui.test.mjs', '--self-test'],
  'yakusoku.test.mjs',                    // ★画面に 書いてある「約束」を 押して 確かめる(字が緑でも 起きない事が在る)
  ['yakusoku.test.mjs', '--self-test'],
  'shift-mado.test.mjs',                  // ★挿入/削除で どちらへ詰めるかを聞く窓(4つ・本当に その向きに 動くか)
  ['shift-mado.test.mjs', '--self-test'],
  'cond-format.test.mjs',                 // ★条件付き書式の当たり判定(実Excelの真値と突き合わせ)
  ['cond-format.test.mjs', '--self-test'],
  'cond-format-ui.test.mjs',              // ★本物の画面で 実際に押す(部品が緑=画面で使える ではない)
  ['cond-format-ui.test.mjs', '--self-test'],
  'ref-graph-forms.test.mjs',             // ★Excelの仕様の側から形を1つずつ（41形＋重い形3つ）
  ['ref-graph-forms.test.mjs', '--self-test'],
  'ref-graph.test.mjs',                   // ★ブック全体の参照の網（別シート401本を落とさない・固まらない）
  ['ref-graph.test.mjs', '--self-test'],
  'book-scan-ui.test.mjs',                // ★調べている間の知らせ（300ms・何枚目・％・必ず消す）
  ['book-scan-ui.test.mjs', '--self-test'],
  'html-script-syntax.test.mjs',          // ★HTMLの中の<script>が 構文で死んでいないか（lint緑・テスト緑でも 死ぬ）
  ['html-script-syntax.test.mjs', '--self-test'],
  'kirikae-egaku.test.mjs',               // ★切り替えた行が「画面に描かれる所」に在るか（選ばれた数ではなく 描かれた物）
  ['kirikae-egaku.test.mjs', '--self-test'],
  'kirikae.test.mjs',                     // ★切り替え＝見る人・見る月を変える（シートは増えない・戻すと元どおり）
  ['kirikae.test.mjs', '--self-test'],
  'vba-tejun.test.mjs',                   // ★マクロから 手順を取り出して レシピにする（取り出せた分だけ・数で言う）
  ['vba-tejun.test.mjs', '--self-test'],
  'vba-ui.test.mjs',                      // ★マクロが 何をしているかを 画面で 出す（描かれた字で 数える）
  ['vba-ui.test.mjs', '--self-test'],
  'vba.test.mjs',                         // ★マクロ（VBA）を 読む・見立てる＝読めた≠正しく読めた・AIには要約だけ
  ['vba.test.mjs', '--self-test'],
  'kikan.test.mjs',                       // ★何月分・◯日からの分（締め期間は Timeally が正本・124通り）
  ['kikan.test.mjs', '--self-test'],
  'kiridashi-jitsubutsu.test.mjs',        // ★司さんの実物で 1人分×何月分を 切り出す（描かれた字で 数え直す）
  ['kiridashi-jitsubutsu.test.mjs', '--self-test'],
  'recipe-ui.test.mjs',                   // ★8-③ 覚えた手順を 画面で押す（AIを呼ぶ前に 機械が当てる・上限を置かない）
  ['recipe-ui.test.mjs', '--self-test'],
  'teian.test.mjs',                       // ★8 提案＝うちから「こう直せますよ」（0円・向こうから出る）
  ['teian.test.mjs', '--self-test'],
  'chizu-horu.test.mjs',                  // ★7 地図＋掘る（地図は数千文字・掘りっぱなしにしない・金額はAIに書かせない）
  ['chizu-horu.test.mjs', '--self-test'],
  'recipe.test.mjs',                      // ★6 レシピ＝手順を覚えて 2回目からはAIを呼ばない（記録の行数で数える）
  ['recipe.test.mjs', '--self-test'],
  'rireki.test.mjs',                      // ★6 履歴＝見る場所（別の入り口から同じに見える・客のブックにタブ0件）
  ['rireki.test.mjs', '--self-test'],
  'chuki.test.mjs',                       // ★注記を外してから読む（見張りの共通部品・同じ型を3回踏んだので決まりにした）
  ['chuki.test.mjs', '--self-test'],
  'ai-genkai.test.mjs',                   // ★掘る数と 叩ける数の 噛み合わせ（2026-09-05）
  ['ai-genkai.test.mjs', '--self-test'],  //   ＝15回叩いて10回で止まり 掘りを全部捨てていた
  'eol.test.mjs',                         // ★行の終わりが CRLF に なると わざと壊す試験が 掴めず
  ['eol.test.mjs', '--self-test'],        //   黙って赤に なる（2026-09-05 に 17本 落ちた）
  'shouko-no-okiba.test.mjs',             // ★消える場所に repo の 物を 置いていないか（2026-09-05）
  ['shouko-no-okiba.test.mjs', '--self-test'], //  ＝同じ日に 2回 落ちた（507個が 消えた／YENの控え）
  'hyou-ireru-ui.test.mjs',               // ★AIの表を入れて★1押しで元に戻る★（本物の画面をjsdomで押す）
  ['hyou-ireru-ui.test.mjs', '--self-test'],// ＝前は1セルずつ積み★16回押さないと戻らなかった★
  'formula-nokori.test.mjs',              // ★2026-09-06に作った14個＝答えは全部★実Excelの実測★
  ['formula-nokori.test.mjs', '--self-test'],
  'formula-kane.test.mjs',                // ★お金の関数22個＝実Excelの答え700本と1本残らず突き合わせ(09-07)
  ['formula-kane.test.mjs', '--self-test'],//  ＝ODDF/ODDLの4個は★合わない形が残るので出していない★
  'formula-yosoku.test.mjs',              // ★予測・統計・単位8個＝実Excelの答え473本と突き合わせ(09-07)
  ['formula-yosoku.test.mjs', '--self-test'],// ＝★CONVERTは「無かった」でなく「4桁に丸めていた」★
  'soto-api.test.mjs',                    // ★外へ出る口が守る物＝許した相手だけ／http・中の網は断る(09-07)
  ['soto-api.test.mjs', '--self-test'],   //  ＝ブラウザから撃つと★4回/4回とも相手に届く★（実測）
  'formula-soto.test.mjs',                // ★外へ出る関数4個＝外は作り物で押す（お金も迷惑もかけない）
  ['formula-soto.test.mjs', '--self-test'],
  'formula-filterxml.test.mjs',           // ★XMLから取り出す＝★XPathはブラウザの物をそのまま使う★(09-07)
  ['formula-filterxml.test.mjs', '--self-test'],// ＝自分で少しだけ作ると★通る式と通らない式が混ざる★
  'nul-nashi.test.mjs',                   // ★生のNULを配る物に入れない＝git が「絵」と決めて差分が消える(09-07)
  ['nul-nashi.test.mjs', '--self-test'],  //  ＝前の見張りは lib/pivot.js ★1本だけ★を見ていた(550本中)
  'formula-areas.test.mjs',               // ★AREAS＝エンジンが式を読む所で断る「とびとび」を通す(09-08 実測89本)
  ['formula-areas.test.mjs','--self-test'],//  ＝間違いが★出さない側に倒れる★ので下限を実測の29本で見張る
  'isref.test.mjs',                       // ★ISREF＝本番で かっこを 壊し 逆の答えを 返していた(09-07 実測25本)
  ['isref.test.mjs', '--self-test'],      //  ＝[^)]+ が入れ子を数えず =FALSE()) を作っていた
  'formula-cell.test.mjs',                // ★CELL＝実Excelに7回 打たせた411本と突き合わせ(09-07)
  ['formula-cell.test.mjs', '--self-test'],//  ＝★" "の中の( は数えない★／$は通貨でない／yyyy/d はG
  'formula-complex.test.mjs',             // ★IM系＝出す字を実Excelと同じ15桁に★丸める★(09-08 実測78本 直る)
  ['formula-complex.test.mjs','--self-test'],// ＝★数を返す4個(IMABS等)に かけると壊れる★ので押して確かめる
  'isomitted.test.mjs',                   // ★ISOMITTED＝『空のマス』と『省かれた引数』は別物(09-08 実Excel実測)
  ['isomitted.test.mjs','--self-test'],   //  ＝空のマスも FALSE／★カンマで省いた形は「まだ」＝黙って逆を返さない★
  'monosashi-mado.test.mjs',                // ★測り道具が「0」を2つの窓で取っているか（09-08 .Value2 が 0でない値に0を返す）
  ['monosashi-mado.test.mjs','--self-test'],// ＝★見せかけの0★と★字の"0"★は .Value2 では見分けが付かない
  'hakaridai-mon.test.mjs',                 // ★測り台が本番と違うのを止める（09-09〜10 に 1日で 5回 転んだ）
  ['hakaridai-mon.test.mjs','--self-test'], // ＝node の台の数は 画面の数ではない／板ごと入れるなら断りを書く
  /* ★1つの式が何マスにも広がる（溢れ）＝SORT/UNIQUE/FILTER/SEQUENCE/LINEST がまとめて動く（09-10）
     ＝エンジンは元から溢れていて★画面がそれを読む道だけが無かった★ */
  /* ★うちが書いた物が実Excelで開けるか（09-10 司さん「ExallyでもExcelでも使えるように確かめながらやってるか」）
     ＝今まで★実Excel→うち★しか見ていなかった。★うち→Excel★を足した */
  'oufuku.test.mjs',
  ['oufuku.test.mjs','--self-test'],
  'afureru.test.mjs',
  ['afureru.test.mjs','--self-test'],
  'shirase-omoku-nuru-na.test.mjs',        // ★知らせの箱を濃い色で塗らない（09-10 司さん「絶対これがええ」）
  ['shirase-omoku-nuru-na.test.mjs','--self-test'], // ＝注意の知らせは08-25に直して在り普通の知らせだけ取り残されていた
  'shisuu-mitame.test.mjs',                 // ★画面に出る字を実Excelと同じに（09-10 実測 574本）
  ['shisuu-mitame.test.mjs','--self-test'], // ＝09-09「画面はきれい」は★中の数と比べていた＝間違い★
  'shisuu-utenai.test.mjs',                 // ★指数の字(=1.64E-14)を打つと #ERROR になっていた（09-10）
  ['shisuu-utenai.test.mjs','--self-test'], // ＝大文字Eをエンジンが数と読めない／セル参照(E14)は壊さない
  'marume-mihari.test.mjs',                 // ★答えを黙って丸めている所が増えていないか（09-08 A群4→0）
  ['marume-mihari.test.mjs','--self-test'],
  'mirr-tomeru-na.test.mjs',                // ★MIRRが実Excelの答える所で断っていないか（09-09 全部マイナスで #DIV/0! を返していた）
  ['mirr-tomeru-na.test.mjs','--self-test'],// ＝出来る物を止めない（式は実Excelの定義から自分で書いた）
  'linest-x-nihon.test.mjs',                // ★LINESTのxが2本以上で静かに違う答えを出さない（09-09 実Excel 0.77 vs うち 7.39）
  ['linest-x-nihon.test.mjs','--self-test'],// ＝断る（正しく計算するのは別の直し）／出来ていた1本の形は止めない // ＝★A群かを機械に判じさせない★新しい丸めは全部赤・人がA〜Eに分ける
  'xirr-mon.test.mjs',                      // ★XIRRが実Excelと同じ条件で同じ誤りを返すか（09-09 全部プラスで 375兆 を返していた）
  ['xirr-mon.test.mjs','--self-test'],      // ＝断らない所は断らない（空マス・同じ日付・0を含む は値を返す）
  'gyousuu-gamen-to-enjin.test.mjs',       // ★画面は100万行なのにエンジンは4万行＝★96%が黙って0★(09-08 本番実測)
  ['gyousuu-gamen-to-enjin.test.mjs','--self-test'],// ＝数を揃えるだけでなく★実際に押して0でない事★まで見る
  'js-sou-sekisho.test.mjs',              // ★JS層の関所を通らない段＝★1度も届かない死にコード★(09-08 実測12か所)
  ['js-sou-sekisho.test.mjs','--self-test'],// ＝死にコードは動かないだけでなく★人が数を間違える★(38個と数えていた)
  'hikisuu-mihari.test.mjs',              // ★引数の見張り＝実Excelが断る所でうちも断る(09-08 実測180本 57→6)
  ['hikisuu-mihari.test.mjs','--self-test'],// ＝★一番危ないのは誤りでなく「静かに違う答え」★(REPLACEB 1.5 が axcde)
  'ugokanai-osu.test.mjs',                // ★「動かない」台帳35個を★本物の式で全部押す★(09-07 指示役)
  ['ugokanai-osu.test.mjs', '--self-test'],//  ＝CONVERTが「動かない」棚に居ながら★間違った数を返していた★
  'kansuu-kabaa.test.mjs',                // ★実Excelの関数を どれだけ動かせているか（09-06 実測 519/432）
  ['kansuu-kabaa.test.mjs', '--self-test'],//  ＝08-29の「507個」は一覧を消して★中身の無い数字★になった
  'hoshi-nashi.test.mjs',                 // ★お客さんの画面に出る字に★を使わない（09-06 指示役）
  ['hoshi-nashi.test.mjs', '--self-test'],//   ＝★は私たちの便りの印。客には壊れた字に見える
  'hairanai.test.mjs',                    // ★書き出すファイルに入らない物を「言う」（09-06 実測）
  ['hairanai.test.mjs', '--self-test'],   //   ＝うちで足したシートは消える。黙って消さない
  'hon-no-nakami.test.mjs',               // ★持ち込んだ本に 何が 組まれているかを 数えて 言う（09-21・司さんの ア②）
  ['hon-no-nakami.test.mjs', '--self-test'], //   ＝★「マクロが 無い」と「読めなかった」を 混ぜない★
  'mochikomi-lib-order.test.mjs',         // ★持ち込み画面と book.html で 読む順が ずれていないか（09-21）
  ['mochikomi-lib-order.test.mjs', '--self-test'], //   ＝★同じ順が 2か所に 在る＝片方だけ直すと 黙って壊れる★
  'prompt-file.test.mjs',                 // ★AIの頭は prompt/ から／台帳とずれたら赤（2026-09-05）
  ['prompt-file.test.mjs', '--self-test'],//   ＝手書き22個のうち17個が間違っていた
  'betsumei-zenbu.test.mjs',              // ★打てば動く別名を機械で全部出して台帳と突き合わせる（09-06 指示役）
  ['betsumei-zenbu.test.mjs', '--self-test'],// ＝JISがどの棚にも無く★AIが知らない★状態だった
  'kansuu-tana.test.mjs',                 // ★台帳の棚を engine に押して確かめる（2026-09-05）
  ['kansuu-tana.test.mjs', '--self-test'],//   ＝YENを「動かない」棚に置き AIが客に嘘を言う所だった
  'ban-wo-okuru.test.mjs',                // ★客のExcelの版がAIまで届くか（会議の穴B・2026-09-05）
  ['ban-wo-okuru.test.mjs', '--self-test'],//  ＝版を1度も送っておらず★全員365扱い★だった
  'shindan.test.mjs',                     // ★5 E2診断1本目＝消えた参照が IFERROR で隠れている（実物122本）
  ['shindan.test.mjs', '--self-test'],
  'shindan-ui.test.mjs',                  // ★本物の画面で 知らせ・一覧・場所へ行く＋直した所の控え
  ['shindan-ui.test.mjs', '--self-test'],
  'ai-jiko-dome.test.mjs',                // ★4 事故止め（1分10回・1日100回・40,000字・2万トークン・429の言い方・記録）
  ['ai-jiko-dome.test.mjs', '--self-test'],
  'xlsm-vba.test.mjs',                    // ★VBA入り(.xlsm)＝開ける・VBAは1バイトも触らない・動かさない・言葉で先に言う
  ['xlsm-vba.test.mjs', '--self-test'],
  'login-gate.test.mjs',                  // ★表の画面にもログイン／忘れた人の逃げ道（無いと二度と入れない）
  ['login-gate.test.mjs', '--self-test'],
  'ai-reason.test.mjs',                   // ★AIに繋がらない時の理由と次の一手／★空のセルでAIを呼ばない(お金)★
  ['ai-reason.test.mjs', '--self-test'],
  'no-dark-green.test.mjs',               // ★使わないと決めた濃い緑(コードは直ったのにCLAUDE.mdの色の表が教え続けていた)
  ['no-dark-green.test.mjs', '--self-test'],
  // P1② 版対応 検証ハーネス
  'excel-version.test.mjs', // ★その式が相手のExcelで動くか(Excelに無い23個＝常時 / 版マーカー14個＝版連動)
  ['excel-version.test.mjs', '--self-test'],
  'xlsx-harness/roundtrip.test.mjs',        // 数式入りxlsxの往復(SheetJS・★新関数の _xlfn.)
  'xlsx-harness/bare-form.test.mjs',        // ★「客が最初に書く形」のケースが無い関数を赤にする(R19の再発防止)
  'xlsx-harness/alias.test.mjs',            // ★日本語UI名→本名(JIS→DBCS / YEN→DOLLAR。入口=エンジン/出口=書き出し)
  'xlsx-harness/xlfn-coverage.test.mjs',    // ★書き出す関数名が分類済みか(_xlfn.の付け忘れを止める)
  'xlsx-harness/version-scope.test.mjs',    // ★「版対応はここまで」の記述と実装がズレたら赤
  'xlsx-harness/compare.mjs',               // Excelの真値と突合(新規の不一致があれば赤)
  ['xlsx-harness/compare.mjs', '--self-test'], // ★わざと壊して赤になるかの自己確認
  ['xlsx-harness/nesting-audit.mjs', '--probe', '--check'], // ★入れ子で壊れる式が増えていないか
  /* ★★ 2026-09-16 ･･･ ★下の 5本は どこからも 呼ばれて いませんでした★★
       名簿（この FILES）にも 無く、`.github/workflows/*.yml` にも 無かった
       ⇒★★試験の 顔を して 1回も 走って いませんでした★★
       ★xlfn-morenashi は commit 586d2eb で
         「試験の 登録も 自動で 通りました（登録漏れ 0件）」と 報告しましたが
         ★その 報告は 間違い★でした。
         （`scripts/tests-registered.mjs` は ★playwright の 一覧★だけ 見て おり
           ★この FILES を 見て いません★＝★別の 物を 数えた 緑★）
       ⇒★見張り `tests/souname-meibo.test.mjs` を 置きました★
         （★足す 前に 5本を 名指しで 赤に する 事を 確かめました★）

       ★★ただし `hozon-de-kieru.test.mjs` は ★ここに 入れて は いけません★★
         ★わざと 赤のまま★と 2026-09-07（commit 20ac3f6）に 決めて あります。
         ＝「うちで 足した シートは 書き出すと 消える」を ★実測で 固定した 物★
         訳と 戻す 条件は ★`tests-no-ci.json`★ に 書いて あります。
         ★私は 1度 ここに 入れて 総なめを 赤に しました★（★自分の 覚書きを 読んで いなかった★）
         ⇒★見張り `souname-meibo` に ★`tests-no-ci.json` を 読ませました★★ */
  ['xlfn-morenashi.test.mjs'],       // ★`_xlfn.` の 一覧に 漏れが 無いか（漏れると 実Excel で #NAME?）
  ['yearfrac-basis1.test.mjs'],      // ★YEARFRAC basis 1 が 実Excel と 合うか
  ['workflow-dougu-aru.test.mjs'],   // ★見張りが 呼ぶ 道具が repo に 在るか
  ['souname-meibo.test.mjs'],         // ★tests/ の 試験が 1本 残らず 走って いるか
  ['oddf-kirwake.test.mjs'],          // ODDFPRICE の 切り分けで 分かった 事を 守る
  ['oddf-basho.test.mjs'],            // ★新しい 数え方を どこに 当てるか★（A には 当てない）
  ['junretsu.test.mjs'],              // ★PERMUT ／ PERMUTATIONA★（紙 2本・境目は 未測定）
  ['xmatch.test.mjs'],               // ★XMATCH★（紙 12本・字の型紙は 未測定）
  ['percentrank.test.mjs'],          // ★PERCENTRANK★（紙 6本・有効桁の 切り方は 未測定）
  ['asc-dbcs.test.mjs'],             // ★ASC ／ DBCS★（紙 22行・表は 機械の Unicode から）
  ['lenb.test.mjs'],                 // ★LENB／LEFTB／RIGHTB／MIDB★（紙 12本・外した 1本を 名指し）
  ['textafter.test.mjs'],            // ★TEXTAFTER／TEXTBEFORE★（★逆だった 2本を 一番 上に★）
  ['aggregate.test.mjs'],            // ★AGGREGATE★（★外した 2本＝選択7・機能19 を 一番 上に★）
  ['shiki-wo-osu-webkit.mjs'],       // ★実UIで 式を 打って 読む★（★繋ぐ 前の 姿を 守る★）
  ['kobore-kakidashi-webkit.mjs'],  // ★画面で 溢れを 打って 本番の 道で 書き出す★（★お客さんの 道★・2026-09-20）
  ['doko-wo-sawatta-bun-webkit.mjs'], // ★どこを 触ったか＝1文 ＋ 畳んだ 詳しく★（★畳みは 高さで 数える★・2026-09-22）
  ['oboeru-nakara-banchi.test.mjs'],  // ★覚えた 番地を 誰も 書き換えないか★（59.5秒→32.1秒の 前提・2026-09-25）
  ['oboeru-nakara-banchi.test.mjs', '--self-test'], //   ＝★書き換えると 狂うのは 書き換えた 所では ない★
  ['zero-wo-kakusu-webkit.mjs'],     // ★ゼロを 隠す 板では 0 を 描かない★（実Excel と 違う 9,163個の うち 6,743個・2026-09-25）
  ['hiraita-toki-keisan-webkit.mjs'],  // ★理由が 1つも 無い 本は 開いた 直後に 計算しない★（23.7秒 → 1.7秒・2026-09-25）
  ['mado-no-kazu-to-kaita-kazu-webkit.mjs'],  // ★窓に 出る 数と 本当に 書き込まれた 数が 合うか★（1マス→窓3044の 実物・2026-09-25）
  ['naoshite-hozon-webkit.mjs'],  // ★お客さんが 入力値を 直して 書き出せるか★（記録8 で 1冊 出なかった・2026-09-25）
  ['akimasu-ni-utsu-webkit.mjs'],  // ★空いて いる マスに 打って 書き出せるか★（行ごと 無い 所で 1冊も 出なかった・2026-09-26）
  ['webkit-hashiru.test.mjs'],  // ★どの webkit の 見張りも CI の どこかで 本当に 走るか★（19本中 15本が どこでも 走って いなかった・2026-10-04）
  ['webkit-hashiru.test.mjs', '--self-test'],  //   ＝★webkit.yml から 名簿の 段を 消すと 赤★
  ['kakareta-ji-webkit.mjs'],  //★画面に 描かれた 字（render の fillText）が 実Excel と 同じか★（書式付きの 0 が 描かれて いなかった・2026-10-04）
  ['kakareta-ji-webkit.mjs', '--self-test'],  //   ＝★_数が入らないか を いつも 偽に して 赤に なるか★
  ['wareki-keika-gamen-webkit.mjs'],  //★和暦・経過時間の 書式の マスが 画面で 実Excel と 同じ 字か★（SSF で 48/155 → 台で 155/155・2026-10-04）
  ['hozon-no-ayamari-wo-yomu-webkit.mjs'],  //★答えが 誤りで 保存された 式は 開いた だけで 誤りの 字★（実物 69個 #REF! が 数に 化けた・2026-10-02）
  ['soto-wo-kiku-webkit.mjs'],  //★外へ つながる 式が 在る 本だけ お客さんに 訊く★（司さん 09-25「おすすめで 直せ」・2026-09-26）
  ['yomenakatta-webkit.mjs'],        // ★ファイルを 読めなかったら 読み直す★（OneDrive で 24回中 15回 断られる・2026-09-25）
  ['mikire-webkit.mjs'],             // ★字は 隣が 空なら はみ出す（実Excel と 同じ）★（★絵で 数える★・2026-09-24）
  ['machi-no-kime.test.mjs'],        // ★待ちの 決めが 効いて いるか★（2つ目は 引数・決めは 3つ目／2026-09-24）
  ['machi-no-kime.test.mjs', '--self-test'], //   ＝★30秒で 黙って 落ちる＝遅い本を「開けない本」と 言い間違える★
  ['kinji-ji.test.mjs'],             // ★禁じられた 字が 増えて いないか★（上限の 門）
  ['souname-matome.test.mjs'],       // ★束に 分けた 総なめの まとめが 赤を 出せるか★（途中で 切られた 束を 緑と 見間違えかけた・2026-10-04）
  ['ci-mitodoke.test.mjs'],          // ★CI の 見届けの 道具が 偽の 緑を 返さないか★（scratchpad の 写しで 4回 見届けて いた・2026-10-10）
  ['ci-mitodoke.test.mjs', '--self-test'], //   ＝判じを 19通り 壊すと 赤が 出るか
  ['sakusha-mihon.test.mjs'],        // ★Office ファイルの 作者の 欄に 名前が 戻って いないか★（作り直すと Excel が サインインの 名前を 入れる・2026-10-10）
  ['sakusha-mihon.test.mjs', '--self-test'], //   ＝作り物の 名前を 入れた 写し・読めない zip で 赤に なるか
  ['shiken-no-kuchi.test.mjs'],      // ★実の 倉庫に 触る 道具の 門＝試験用の 口（+タグ が 白名簿）以外を 止めるか★（受信箱の メールを repo の 外へ・2026-10-10）
  ['pre-push-noreply.test.mjs'],     // ★押す 前の 作者の 門＝noreply で ない メールの commit を 止めるか★（司さん「アドレスが分からんようにしろや」・2026-10-11）
  ['bessel.test.mjs'],               // BESSELI/J/K/Y
  ['jitsuexcel-ga-machigai.test.mjs'], // 実Excel が 間違って いるの 逃げ道を 塞ぐ
  ['gomi-file.test.mjs'],            // 書き損じで 出来た ゴミを repo に 置かない
  ['oddf-hasuu-ga-hasuu-denai.test.mjs'], // ★端数が「端数で ない」時は #NUM!
  ['tsunagu-mon.test.mjs'],          // ★台を 画面に 繋ぐ 手を 止める 門（★繋ぐ 前から 赤★）
  ['ita-awase.test.mjs'],           // ★板の 写しの 門（★わざと 1マス ずらして 赤に なるのを 見せます★）
  ['oddf-susumi.test.mjs'],         // ★保留の 4個（ODDF / ODDL）の 進みを 数で 押さえる★
  ['oddf-susumi.test.mjs', '--self-test'],   // ★自分で 壊して 赤に なるか★
  ['oddf-nokori.test.mjs'],         // ★ODDF / ODDL の 残りを 実Excel の 50本で 押す★
  ['phonetic-ramuda.test.mjs'],     // ★PHONETIC と 裸の LAMBDA を 実Excel の 6本で 押す★（2026-09-19）
  ['kobore-zenmasu.test.mjs'],      // ★溢れた先の 全マスを 実Excel の 14本で 押す★（2026-09-19）
  ['kobore-kakidashi.test.mjs'],   // ★溢れを xlsx へ 書き出す時に ref= と cm="1" を 運べて いるか★（2026-09-20）
  ['kobore-kakidashi.test.mjs', '--self-test'],   // ★印を 消すと 赤に なるか★
  ['kobore-fusagi.test.mjs'],      // ★塞がれた 溢れ(#SPILL!)を 書き出す時★（★付けないと 実Excel で 黙って 1 に なる★・2026-09-20）
  ['kobore-fusagi.test.mjs', '--self-test'],   // ★#SPILL! 以外に 付けたら 赤か★
  ['kobore-nakami.test.mjs'],       // ★溢れた 先の ★中身★を 実Excel の 紙と 突き合わせる★（2026-09-21）
  ['kobore-nakami.test.mjs', '--self-test'],   // ★紙の 中身を 変えると 赤か★
  ['kobore-basho.test.mjs'],        // ★場所を 見る 関数と ROW/COLUMN の 向き★（2026-09-21）
  ['kobore-basho.test.mjs', '--self-test'],    // ★A2 を 式で 無くすと 変わるか★
  ['xlsx-kazari.test.mjs'],         // ★マスの 飾り（太字・字の色・塗り・罫線）を 自前で 読む★（2026-09-21）
  ['xlsx-kazari.test.mjs', '--self-test'],     // ★styles.xml を 5通り 壊して 赤に なるか★
  ['xlsx-zukei.test.mjs'],          // ★図形（判子）を 自前で 読む★（2026-09-21）
  ['xlsx-zukei.test.mjs', '--self-test'],      // ★drawing1.xml を 5通り 壊して 赤に なるか★
  ['zip-tasu.test.mjs'],            // ★包みに 部品を 足せるか★（ア＝全部 保存する・2026-09-21）
  ['zip-tasu.test.mjs', '--self-test'],        // ★足さない/消す/書き換えるで 数が 動くか★
  ['ita-wo-tasu.test.mjs'],         // ★足した 板が 書き出す 先に 入るか★（ア・2026-09-21）
  ['ita-wo-tasu.test.mjs', '--self-test'],     // ★足さない/型を捨てる/中身を変える★
  ['kansuu46-1taba.test.mjs'],       // ★台に 無かった 49個の 1束目（7個）を 紙で 押す
  ['kansuu46-9wakume.test.mjs'],     // ★9枠目の 紙で 押す（★紙と 同じ commit で 出す★）
  ['kansuu46-8wakume.test.mjs'],     // ★8枠目の 紙で 押す（★紙は 在ったが 押されて いなかった★）
  /* ★経営者1 の 1押し★（★私が `git add -A` で 先に commit して しまった★）
     ⇒★名簿に 足さないと 見張りが 赤の まま★なので ここで 足します
     ⇒★中身は 1文字も 触って いません★ */
  ['ramuda-honban.test.mjs']         // ★ラムダの 一族を お客さんの 道で 押す（経営者1）
];

/* ★直に叩いた時だけ 走る★（2026-08-29）
   ここを 守らないと 一覧を 読むだけの道具が ★全部の試験を 走らせてしまう★（実際に 起きた）。 */
if (require.main !== module) { module.exports = { FILES }; return; }

let ng = 0;
const 失敗一覧 = [];
for (const f of FILES) {
  const [file, ...args] = Array.isArray(f) ? f : [f];
  console.log('\n=== ' + file + (args.length ? ' ' + args.join(' ') : '') + ' ===');
  /* ★落ちた理由を必ず出す（2026-08-23）★
     CIで1回 赤になったのに ★同じコミットを回し直したら緑★＝★ムラ★だった。
     その時 出ていたのは「★ N ファイルで失敗」だけで、
     ★中で死んだのか（signal）／自分で1を返したのか（status）が 分からなかった★。
     ⇒ ★どちらかを 必ず1行 出す★（次に赤くなった時、新しい壊れ か ムラ かを その場で見分ける） */
  /* ★中で殺される(SIGTRAP)のを止める（2026-08-24）★
     CIで excel-parity が ★signal=SIGTRAP★ で死んだ＝V8 が力尽きた（本物の壊れではない）。
     本物の画面(book.html)を jsdom に丸ごと載せる検査は 重い。★積める量を増やす★。
     ⇒ 落ちた理由を出す1行が無ければ ★「新しい壊れ」と区別できなかった★ */
  try { execFileSync(process.execPath, ['--max-old-space-size=4096', path.join(__dirname, file), ...args], { stdio: 'inherit' }); }
  catch (e) {
    ng++;
    const 印 = e && e.signal ? ('★中で殺された(signal=' + e.signal + ')★＝新しい壊れではない可能性')
      : ('自分で ' + (e && e.status !== undefined && e.status !== null ? e.status : '?') + ' を返した');
    console.log('  ★落ちた★ ' + file + (args.length ? ' ' + args.join(' ') : '') + ' ･･･ ' + 印);
    失敗一覧.push(file + (args.length ? ' ' + args.join(' ') : '') + '（' + 印 + '）');
  }
}
console.log('\n' + (ng ? '★ ' + ng + ' ファイルで失敗' : '全テストファイル 緑'));
for (const 名 of 失敗一覧) console.log('   ・' + 名);
process.exit(ng ? 1 : 0);
