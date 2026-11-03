/* ribbon-actions.js — ★リボンの ボタン1つ1つの 働き★ 2026-08-29
 *
 *  ★なぜ 間に 1枚 挟むか（★実際に 踏んだ★）★
 *    最初 リボンから 画面の関数を ★直接★ 呼ぶ形に した。すると:
 *      applyFormat は 引数が ★2つ★（key, value）なのに 1つで 呼んでいた
 *      openColorPalette は ★event★ を 取るのに 文字を 渡していた
 *    ⇒ ★押しても 何も起きないボタン★が 出来る所だった
 *      （「行き先の関数が 在るか」だけ 見ていて 気づけなかった）
 *    ⇒ ★呼び方を ここに 1か所で 書く★。試験は ★実際に 押して 画面が 変わったか★ を 見る。
 *
 *  ★決まり★
 *    ・1つの働き＝1つの関数（引数は 取らない）
 *    ・画面の関数が 無い時は ★何もしない★（落ちない）
 *    ・★出来ていない物は ここに 書かない★＝リボンに ボタンが 出ない
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.RibbonActions = factory();
}(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  function W() { return (typeof window !== 'undefined') ? window : {}; }
  function 呼ぶ(名) {
    var f = W()[名];
    if (typeof f !== 'function') return false;
    var 引数 = Array.prototype.slice.call(arguments, 1);
    f.apply(null, 引数);
    return true;
  }

  var A = {
    /* ── クリップボード ── */
    貼り付け:   function () { return 呼ぶ('ctxPaste'); },
    切り取り:   function () { return 呼ぶ('ctxCut'); },
    コピー:     function () { return 呼ぶ('ctxCopy'); },
    /* ★書式のコピー／貼り付け★＝中身は 写さない（実Excelと 同じ） */
    書式のコピー: function () { return 呼ぶ('書式をコピー'); },

    /* ── フォント ── */
    太字:       function () { return 呼ぶ('toggleFormat', 'bold'); },
    斜体:       function () { return 呼ぶ('toggleFormat', 'italic'); },
    下線:       function () { return 呼ぶ('toggleFormat', 'underline'); },
    取り消し線: function () { return 呼ぶ('toggleFormat', 'strike'); },
    罫線:       function () { return 呼ぶ('applyBorderAll'); },
    下罫線:     function () { return 呼ぶ('applyFormat', 'borderBottom', true); },
    塗りつぶしの色: function () { return 呼ぶ('openColorPalette', null, 'bg'); },
    フォントの色:   function () { return 呼ぶ('openColorPalette', null, 'fg'); },
    黄色で塗る:     function () { return 呼ぶ('applyFormat', 'bg', '#FFFF00'); },
    赤い字にする:   function () { return 呼ぶ('applyFormat', 'color', '#FF0000'); },
    フォントの設定: function () { return 呼ぶ('openFmtModal'); },
    /* サイズは 画面の 入力欄を 動かす（1か所で 持つ） */
    /* ★引き取る物★＝押す働きは 元の物が 持っている。名前だけ 要る。 */
    字の大きさ: function () { return true; },
    字を大きく: function () { return _サイズ(1); },
    字を小さく: function () { return _サイズ(-1); },

    /* ── 配置 ── */
    左揃え:     function () { return 呼ぶ('applyFormat', 'align', 'left'); },
    中央揃え:   function () { return 呼ぶ('applyFormat', 'align', 'center'); },
    右揃え:     function () { return 呼ぶ('applyFormat', 'align', 'right'); },
    上揃え:     function () { return 呼ぶ('applyFormat', 'valign', 'top'); },
    上下中央:   function () { return 呼ぶ('applyFormat', 'valign', 'middle'); },
    下揃え:     function () { return 呼ぶ('applyFormat', 'valign', 'bottom'); },
    折り返す:   function () { return 呼ぶ('applyFormat', 'wrap', true); },
    結合して中央: function () { return 呼ぶ('applyMergeCells'); },
    字下げを増やす: function () { return 呼ぶ('applyIndent', 1); },
    字下げを減らす: function () { return 呼ぶ('applyIndent', -1); },
    配置の設定: function () { return 呼ぶ('openFmtModal'); },

    /* ── 数値 ── */
    書式の設定:     function () { return 呼ぶ('openFmtModal'); },
    /* ★Excelの ボタンは 2つ（増やす／減らす）★＝1つで 回る形は Excelと 違う */
    小数を増やす:   function () { return 呼ぶ('小数を増やす'); },
    小数を減らす:   function () { return 呼ぶ('小数を減らす'); },
    パーセント:     function () { return 呼ぶ('applyFormat', 'numFmt', '0%'); },
    桁区切り:       function () { return 呼ぶ('applyFormat', 'numFmt', '#,##0'); },
    通貨:           function () { return 呼ぶ('applyFormat', 'numFmt', '¥#,##0'); },

    /* ── スタイル ── */
    条件付き書式:   function () { return 呼ぶ('openCondFormat'); },
    /* ★セルのスタイル★＝中身は 実Excelから 読んだ 真値（見出し1/2/3・良い・悪い…） */
    セルのスタイル: function () { return 呼ぶ('セルのスタイルを開く'); },

    /* ── セル ── */
    行を挿入:   function () { return 呼ぶ('ctxInsertRow'); },
    行を削除:   function () { return 呼ぶ('ctxDeleteRow'); },
    列を挿入:   function () { return 呼ぶ('ctxInsertCol'); },
    列を削除:   function () { return 呼ぶ('ctxDeleteCol'); },
    /* ★セルだけ ずらす★（行や列を まるごと 動かす物とは 別）
       実Excel実測＝★選んだ セルの 列だけ 動く（隣の列は 動かない）★ */
    セルを挿入: function () { return 呼ぶ('セルを下へ挿入'); },
    セルを削除: function () { return 呼ぶ('セルを上へ削除'); },
    /* ★行の高さ／列の幅★＝実Excel実測（既定 17.7pt / 54pt） */
    行の高さ:   function () { return 呼ぶ('行の高さを開く'); },
    列の幅:     function () { return 呼ぶ('列の幅を開く'); },

    /* ── 編集 ── */
    検索と選択: function () { return 呼ぶ('findNextCell'); },
    元に戻す:   function () { return 呼ぶ('doUndo'); },
    やり直す:   function () { return 呼ぶ('doRedo'); },
    オートSUM:  function () { return 呼ぶ('autoSum'); },
    /* ★フィル★＝Excelの「フィル ▼」。うちは 上から／左から の 2つ（Ctrl+D / Ctrl+R と 同じ物） */
    上から写す: function () { return 呼ぶ('fillFromEdge', 'down'); },
    左から写す: function () { return 呼ぶ('fillFromEdge', 'right'); },
    /* ★クリア★＝実Excelで 測った 3通り（中身だけ／書式だけ／すべて） */
    中身を消す: function () { return 呼ぶ('中身を消す'); },
    書式を消す: function () { return 呼ぶ('書式を消す'); },
    すべて消す: function () { return 呼ぶ('すべて消す'); },

    /* ── 数式 ── */
    関数の挿入: function () { return 呼ぶ('openFnPalette'); },
    /* ★関数ライブラリの 分類★＝Excelの 財務／論理／文字列／日付／検索／数学 に 当てる。
       うちの分類（FN_CATS）に 無い物（財務・エンジニアリング…）は ★作っていないので 結ばない★ */
    論理の関数:     function () { return 呼ぶ('openFnPalette', '論理'); },
    文字列の関数:   function () { return 呼ぶ('openFnPalette', '文字列'); },
    日付の関数:     function () { return 呼ぶ('openFnPalette', '日付'); },
    検索の関数:     function () { return 呼ぶ('openFnPalette', '検索'); },
    数学の関数:     function () { return 呼ぶ('openFnPalette', '数学'); },
    その他の関数:   function () { return 呼ぶ('openFnPalette', 'その他'); },

    /* ── データ ── */
    昇順:       function () { return 呼ぶ('sortRange', 'asc'); },
    降順:       function () { return 呼ぶ('sortRange', 'desc'); },
    絞り込む:   function () { return 呼ぶ('filterByValue'); },
    絞り込みを解除: function () { return 呼ぶ('clearFilter'); },
    /* ★2つの鍵で 並べ替え★＝実Excel実測（同じ値は 2つ目の鍵で 決まる） */
    並べ替え:       function () { return 呼ぶ('並べ替えの窓を開く'); },
    入力の決まり:   function () { return 呼ぶ('openValid'); },

    /* ── 数式（ワークシート分析・計算方法）── */
    数式の表示:       function () { return 呼ぶ('数式の表示を切り替える'); },
    すべて再計算:     function () { return 呼ぶ('すべて再計算'); },
    このシートを再計算: function () { return 呼ぶ('このシートを再計算'); },
    エラーチェック:   function () { return 呼ぶ('openShindan'); },

    /* ── 数式（定義された名前）── */
    名前の定義:       function () { return 呼ぶ('名前の窓を開く'); },
    名前の管理:       function () { return 呼ぶ('名前の窓を開く'); },

    /* ── データ ── */
    重複の削除:       function () { return 呼ぶ('重複を削除'); },
    区切り位置:       function () { return 呼ぶ('区切り位置を開く'); },

    /* ── 自動化（Excelは Office スクリプト。うちは ★レシピ★）── */
    /* ★レシピ★＝1回目で 手順を 覚え、2回目からは AIを 呼ばない（うちの 答え）
       Excelの Office スクリプト は Microsoftのクラウドで 動く 別物＝そちらは 対象外 */
    レシピ:       function () { return 呼ぶ('openMacro'); },
    履歴:         function () { return 呼ぶ('openRireki'); },
    行を再表示:   function () { 呼ぶ('ctxUnhideRow'); return 呼ぶ('ctxUnhideCol'); },

    /* ── 表示 ── */
    枠を固定:   function () { return 呼ぶ('freezePanes'); },
    固定をやめる: function () { return 呼ぶ('unfreezePanes'); },
    ズーム100:  function () { return 呼ぶ('ズーム100'); },
    /* ★画面の枠線★＝紙に刷る枠線とは 別物（Excelも 別々に 持っている） */
    枠線を表示: function () { return 呼ぶ('枠線の表示を切り替える'); },
    見出しを表示: function () { return 呼ぶ('見出しを出すか'); },
    数式バー:     function () { return 呼ぶ('数式バーを出すか'); },
    /* ★表示｜表示★（2026-08-30 測り直しで 見つけた） */
    ルーラー:     function () { return 呼ぶ('ルーラーを出すか'); },
    ナビゲーション: function () { return 呼ぶ('ナビゲーションを開く'); },
    セルにフォーカス: function () { return 呼ぶ('セルにフォーカスを切り替える'); },
    マクロ:     function () { return 呼ぶ('openMacro'); },
    /* ★[開発]→[Visual Basic]＝マクロの 字（コード）を そのまま 見る★ 2026-09-03
       ★[マクロ]（何をしているかの 見立て）とは 別物★＝あちらは 見立て、こちらは ★字そのもの★。
       ★読むだけ・動かしません★（うちは VBA を 実行しない＝司さんの 決め）。
       前は Visual Basic も `マクロ` に つないでいた
       ＝★別の名前の ボタンが 同じ窓を 開いていた★
       （2026-09-03「同じ組に 同じ働きが 2つ」の 見張りが 見つけた）。 */
    VBAの字を見る: function () { return 呼ぶ('VBAの字を見る'); },

    /* ── 校閲（保護）── */
    /* ★実Excel実測＝新しいセルは 全部 ロック／ロックは 保護して 初めて 効く★
       ★合言葉は 掛けない★（忘れたら 二度と 開けない／Excelの 合言葉も 守りとしては 弱い） */
    シートの保護:   function () { return 呼ぶ('シートの保護を切り替える'); },
    セルのロック:   function () { return 呼ぶ('選んだ所のロックを切り替える'); },

    /* ── 校閲（コメント）── */
    /* ★実Excel実測＝セルに 付く／作者が 付く／既定では 見えない（赤い印だけ）★ */
    新しいコメント: function () { return 呼ぶ('コメントの窓を開く'); },
    コメントを消す: function () { return 呼ぶ('コメントを消す'); },
    前のコメント:   function () { return 呼ぶ('前のコメントへ'); },
    次のコメント:   function () { return 呼ぶ('次のコメントへ'); },
    コメントの表示: function () { return 呼ぶ('コメントの一覧'); },

    /* ── 挿入（リンク）── */
    /* ★実Excel実測＝字は #467886 の 下線つき★（同じ書類を Excelで 開いた時に 同じに 見える） */
    リンク: function () { return 呼ぶ('リンクの窓を開く'); },

    /* ── 挿入（グラフ）── */
    /* ★実Excel実測＝既定は 集合縦棒／360×216／凡例は 系列2本以上／タイトルは 出す★
       うちに 無い形（階層・統計・散布図・マップ…）は ★結ばない＝ボタンを 出さない★ */
    縦棒グラフ:   function () { return 呼ぶ('グラフを作る', 'column'); },
    折れ線グラフ: function () { return 呼ぶ('グラフを作る', 'line'); },
    円グラフ:     function () { return 呼ぶ('グラフを作る', 'pie'); },
    /* ★種類ごとの 既定（凡例の 出る/出ない・軸の 数）は 実Excelを 測って 合わせた★
       … lib/chart.js の 見出し と tests/fixtures/excel-chart-types.txt */
    おすすめグラフ:     function () { return 呼ぶ('おすすめグラフ'); },
    散布図:             function () { return 呼ぶ('グラフを作る', 'scatter'); },
    階層構造グラフ:     function () { return 呼ぶ('グラフを作る', 'treemap'); },
    統計グラフ:         function () { return 呼ぶ('グラフを作る', 'histogram'); },
    ウォーターフォール: function () { return 呼ぶ('グラフを作る', 'waterfall'); },
    複合グラフ:         function () { return 呼ぶ('グラフを作る', 'combo'); },
    /* ── 挿入（スパークライン）── ★実Excel実測＝飾りなし・勝敗だけマイナス別・縦軸は1本ごと★ */
    スパークライン折れ線: function () { return 呼ぶ('スパークラインの窓を開く', 'line'); },
    スパークライン縦棒:   function () { return 呼ぶ('スパークラインの窓を開く', 'column'); },
    スパークライン勝敗:   function () { return 呼ぶ('スパークラインの窓を開く', 'winloss'); },

    /* ── 挿入（テーブル）── ★実Excel実測＝名前は テーブル1・見出し行あり・しま（行）あり★ */
    テーブル: function () { return 呼ぶ('表の窓を開く'); },

    /* ── 挿入（記号と特殊文字）── ★実Excelの 並びは COMから 読めず 未測定＝うちの決め（理由つき）★ */
    記号と特殊文字: function () { return 呼ぶ('記号の窓を開く'); },

    /* ── 挿入（テーブル→フォーム）── ★実測＝実Excelは まわりの塊（CurrentRegion）を 相手に する★ */
    フォーム: function () { return 呼ぶ('フォームの窓を開く'); },

    /* ── 数式 ── ★実Excel実測＝計算方法の既定は 自動（-4105）★ */
    計算方法の設定:     function () { return 呼ぶ('計算方法の窓を開く'); },
    最近使った関数:     function () { return 呼ぶ('関数の窓を開く', '最近'); },
    財務:               function () { return 呼ぶ('関数の窓を開く', '財務'); },
    数式で使用:         function () { return 呼ぶ('数式で使用の窓を開く'); },
    選択範囲から作成:   function () { return 呼ぶ('範囲から名前の窓を開く'); },

    /* ── 数式（ワークシート分析）── ★実測＝ShowPrecedents で 矢印2本・ClearArrows で 0★ */
    参照元のトレース:   function () { return 呼ぶ('参照元のトレース'); },
    参照先のトレース:   function () { return 呼ぶ('参照先のトレース'); },
    トレース矢印の削除: function () { return 呼ぶ('トレース矢印の削除'); },
    数式の検証:         function () { return 呼ぶ('数式の検証を開く'); },
    ウォッチウィンドウ: function () { return 呼ぶ('ウォッチを開く'); },

    /* ── 表示（ズーム）── ★実Excel実測＝10％〜400％／既定 100★ */
    ズーム:                 function () { return 呼ぶ('ズームの窓を開く'); },
    選択範囲に合わせて拡大縮小: function () { return 呼ぶ('選択範囲に合わせる'); },

    /* ── 校閲 ── ★実測＝守ると シートが 足せない／許可範囲は 名前＋範囲★ */
    ブックの保護:           function () { return 呼ぶ('ブックの保護を開く'); },
    範囲の編集を許可する:   function () { return 呼ぶ('範囲の編集を開く'); },
    ブックの統計情報:       function () { return 呼ぶ('ブックの数を開く'); },

    /* ── データ ── ★実測＝FlashFill は 1つの 見本から 覚える／再適用は ApplyFilter★ */
    フラッシュフィル: function () { return 呼ぶ('フラッシュフィル'); },
    再適用:           function () { return 呼ぶ('絞りを再適用'); },
    WhatIf分析:       function () { return 呼ぶ('ゴールシークを開く'); },

    /* ── ページ レイアウト ── ★実測＝余白 上下1.905cm/左右1.778cm・用紙A4・印刷範囲は空★ */
    余白:                 function () { return 呼ぶ('ページの詳しい設定を開く', '余白'); },
    サイズ:               function () { return 呼ぶ('ページの詳しい設定を開く', '用紙'); },
    印刷範囲:             function () { return 呼ぶ('ページの詳しい設定を開く', '印刷範囲'); },
    シートのオプション表示: function () { return 呼ぶ('ページの詳しい設定を開く', '表示'); },
    シートのオプション印刷: function () { return 呼ぶ('ページの詳しい設定を開く', '印刷'); },

    /* ── ホーム ── ★実測＝既定フォントは 游ゴシック 11／向きの数は Excelと 同じ★ */
    フォント:                 function () { return 呼ぶ('フォントの窓を開く'); },
    方向:                     function () { return 呼ぶ('文字の向きの窓を開く'); },
    テーブルとして書式設定:   function () { return 呼ぶ('表の窓を開く'); },
    Officeクリップボード:     function () { return 呼ぶ('クリップボードを開く'); },

    /* ── 挿入（ピボット）── ★実測＝左上「合計 / 金」・総計は 行も 列も 出す★ */
    ピボットテーブル:         function () { return 呼ぶ('ピボットの窓を開く'); },
    おすすめピボットテーブル: function () { return 呼ぶ('ピボットの窓を開く', 'おすすめ'); },
    ピボットグラフ:           function () { return 呼ぶ('ピボットグラフを作る'); },

    /* ── データ（統合・アウトライン）／挿入（スライサー）── */
    統合:       function () { return 呼ぶ('統合の窓を開く'); },
    アウトライン: function () { return 呼ぶ('アウトラインの窓を開く'); },
    スライサー: function () { return 呼ぶ('スライサーを開く'); },

    /* ── データ（予測）／校閲（メモ・パフォーマンス）── */
    予測シート:             function () { return 呼ぶ('予測の窓を開く'); },
    /* ★メモ★＝実Excelの 旧コメント。うちの コメントと 同じ 物（作者つき・赤い印）＝
       同じ窓を 開く（★同じ物を 2つ 持たない★） */
    メモ:                   function () { return 呼ぶ('コメントの窓を開く'); },
    パフォーマンスをチェック: function () { return 呼ぶ('重さを調べる'); },

    /* ── 自動化 ── ★Microsoftの クラウドの 仕組みは 借りないが 中身は うちで 作った★ */
    サブテーブルを作成:   function () { return 呼ぶ('サブテーブルを作る'); },
    ハイパーリンクを削除: function () { return 呼ぶ('リンクを全部消す'); },
    空の行数をカウント:   function () { return 呼ぶ('空の行を数える'); },
    テーブルをJSON:       function () { return 呼ぶ('テーブルをJSONにする'); },
    テーブルからピボット: function () { return 呼ぶ('表からピボット'); },

    /* ── データ（取得と変換）── ★出来る 読み口だけ 並べる★ */
    データの取得:       function () { return 呼ぶ('データの取得を開く'); },
    テキストまたはCSVから: function () { return 呼ぶ('CSVの窓を開く'); },
    最近使ったソース:   function () { return 呼ぶ('openRireki'); },

    /* ── 挿入（図・テキスト）／ページレイアウト（配置）── ★色は 写さず うちの緑★ */
    画像:               function () { return 呼ぶ('画像の窓を開く'); },
    図形:               function () { return 呼ぶ('図形の窓を開く'); },
    テキスト:           function () { return 呼ぶ('テキストを置く'); },
    前面へ移動:         function () { return 呼ぶ('前面へ移動'); },
    背面へ移動:         function () { return 呼ぶ('背面へ移動'); },
    オブジェクトの選択と表示: function () { return 呼ぶ('物の一覧を開く'); },

    /* ── 描画 ── ★ペンの 色と 太さは 実Excelの 項目名の とおり★ */
    描画ツール:     function () { return 呼ぶ('描画を始める'); },
    インクの選択:   function () { return 呼ぶ('選ぶにする'); },
    なげなわ選択:   function () { return 呼ぶ('なげなわにする'); },
    消しゴム:       function () { return 呼ぶ('消しゴムにする'); },
    ペン黒:         function () { return 呼ぶ('ペンにする', 0); },
    ペン赤:         function () { return 呼ぶ('ペンにする', 1); },
    鉛筆:           function () { return 呼ぶ('ペンにする', 2); },
    蛍光ペン:       function () { return 呼ぶ('ペンにする', 3); },
    万年筆:         function () { return 呼ぶ('ペンにする', 4); },
    ブラシペン:     function () { return 呼ぶ('ペンにする', 5); },
    ペンを追加:     function () { return 呼ぶ('ペンを足す'); },
    インクの再生:   function () { return 呼ぶ('インクを再生'); },
    インクのヘルプ: function () { return 呼ぶ('インクのヘルプ'); },

    /* ── ページ レイアウト（拡大縮小印刷・配置）── ★実測＝既定 100％・横1・縦1／10〜400★ */
    拡大縮小印刷:   function () { return 呼ぶ('拡大縮小を開く'); },
    横ページ:       function () { return 呼ぶ('拡大縮小を開く'); },
    縦ページ:       function () { return 呼ぶ('拡大縮小を開く'); },
    倍率を増やす:   function () { return 呼ぶ('印刷倍率を足す', 5); },
    倍率を減らす:   function () { return 呼ぶ('印刷倍率を足す', -5); },
    ページ設定ボタン: function () { return 呼ぶ('ページ設定を開く'); },
    配置そろえる:   function () { return 呼ぶ('配置の窓を開く'); },
    グループ化:     function () { return 呼ぶ('グループの窓を開く'); },
    回転:           function () { return 呼ぶ('回転の窓を開く'); },

    /* ── 表示（シート ビュー・ブックの表示・ウィンドウ）── ★実測＝View 1/2/3・分割は Panes 4★ */
    シートビュー切替:   function () { return 呼ぶ('シートビューの窓を開く'); },
    シートビュー保持:   function () { return 呼ぶ('シートビューを保持'); },
    シートビュー終了:   function () { return 呼ぶ('シートビューを終了'); },
    シートビュー新規:   function () { return 呼ぶ('シートビューを新規'); },
    シートビュー設定:   function () { return 呼ぶ('シートビューのオプション'); },
    標準の表示:         function () { return 呼ぶ('標準の表示にする'); },
    改ページプレビュー: function () { return 呼ぶ('改ページプレビューにする'); },
    ページレイアウト表示: function () { return 呼ぶ('ページレイアウト表示にする'); },
    ユーザー設定のビュー: function () { return 呼ぶ('ブックのビューの窓を開く'); },
    新しいウィンドウ:   function () { return 呼ぶ('新しいウィンドウを開く'); },
    ウィンドウ整列:     function () { return 呼ぶ('整列の窓を開く'); },
    ウィンドウ分割:     function () { return 呼ぶ('分割する'); },
    ウィンドウ非表示:   function () { return 呼ぶ('窓を表示しない'); },
    ウィンドウ再表示:   function () { return 呼ぶ('窓を再表示'); },
    並べて比較:         function () { return 呼ぶ('並べて比較'); },
    同時にスクロール:   function () { return 呼ぶ('同時にスクロールを切り替え'); },
    ウィンドウ位置戻す: function () { return 呼ぶ('窓の位置を元に戻す'); },
    ウィンドウ切替:     function () { return 呼ぶ('窓の切り替えを開く'); },
    改ページ:           function () { return 呼ぶ('改ページの窓を開く'); },

    /* ── テーマ・背景・ふりがな ── ★実測＝12の役割／濃淡は0〜240で切り捨て（15/15 一致）★ */
    テーマ:         function () { return 呼ぶ('テーマの窓を開く'); },
    配色:           function () { return 呼ぶ('配色の窓を開く'); },
    テーマのフォント: function () { return 呼ぶ('テーマのフォントの窓を開く'); },
    テーマの効果:   function () { return 呼ぶ('テーマの効果の窓を開く'); },
    背景:           function () { return 呼ぶ('背景の窓を開く'); },
    ふりがな:       function () { return 呼ぶ('ふりがなを入れる'); },

    /* ── 挿入（図解・アイコン・数式・チェックボックス・画面を撮る）── */
    /*    ★実測＝SmartArt は 型159・節5／チェックは セルの TRUE/FALSE／
     *    アイコンは COM から 一覧が 出せない（うちで 描く）★ */
    SmartArt:       function () { return 呼ぶ('図解の窓を開く'); },
    アイコン:       function () { return 呼ぶ('アイコンの窓を開く'); },
    数式を挿入:     function () { return 呼ぶ('数式の窓を開く'); },
    チェックボックス: function () { return 呼ぶ('チェックボックスを入れる'); },
    スクリーンショット: function () { return 呼ぶ('画面を撮って貼る'); },

    /* ── データ（詳細設定・クエリと接続）── ★実測＝見出しは いつも 付く／
     *    条件あり3行・条件あり+重複除く2行・条件なし+重複除く4行・条件なし5行★ */
    詳細設定:       function () { return 呼ぶ('詳細設定を開く'); },
    範囲から:       function () { return 呼ぶ('表から接続を作る'); },
    すべて更新:     function () { return 呼ぶ('すべて更新'); },
    クエリと接続:   function () { return 呼ぶ('接続の窓を開く'); },
    接続のプロパティ: function () { return 呼ぶ('接続のプロパティ'); },
    ブックのリンク: function () { return 呼ぶ('ブックのリンクを見る'); },
    既存の接続:     function () { return 呼ぶ('接続の窓を開く'); },

    /* ── 校閲（文章校正・見やすさ・翻訳・変更内容・共有・インク）── */
    /*    ★実測＝recieve/teh は False・日本語は いつも True／代替テキストは はじめ 空／
     *    共有していないと 変更内容は 断られる（0x800A03EC）★ */
    スペルチェック: function () { return 呼ぶ('スペルを見る'); },
    類義語辞典:     function () { return 呼ぶ('類義語を聞く'); },
    見やすさ検査:   function () { return 呼ぶ('見やすさを見る'); },
    翻訳:           function () { return 呼ぶ('翻訳を聞く'); },
    変更内容を表示: function () { return 呼ぶ('変更内容を見る'); },
    共有を解除:     function () { return 呼ぶ('共有を解除'); },
    インクを非表示: function () { return 呼ぶ('インクを隠す切り替え'); },

    /* ── 開発（XML・コントロール）── ★実測＝結ぶ前は 読み込めない／出せない（IsExportable=False）★ */
    XMLソース:       function () { return 呼ぶ('XMLの窓を開く'); },
    XML対応付けの決め: function () { return 呼ぶ('XMLの決めを見る'); },
    XML拡張パック:   function () { return 呼ぶ('XML拡張パック'); },
    XMLデータの更新: function () { return 呼ぶ('XMLを更新'); },
    XMLインポート:   function () { return 呼ぶ('XMLの窓を開く'); },
    XMLエクスポート: function () { return 呼ぶ('XMLを書き出す'); },
    コントロール挿入: function () { return 呼ぶ('コントロールを挿入'); },
    コントロールの決め: function () { return 呼ぶ('コントロールのプロパティ'); },

    /* ── 描画（変換）・挿入（マップ・タイムライン）── */
    /*    ★実測＝変換の 中の 動きは 見えない／地図の 境目は 向こうの 物／
     *    タイムラインは 日付の 列が 要る★ */
    インクを図形に:   function () { return 呼ぶ('インクを図形に'); },
    インクを数式に:   function () { return 呼ぶ('インクを数式に'); },
    アクションペン:   function () { return 呼ぶ('アクションペン'); },
    マップ:           function () { return 呼ぶ('マップを作る'); },
    タイムライン:     function () { return 呼ぶ('タイムラインを作る'); },

    /* ── 3Dモデル・Webから・データモデル・データ分析 ── */
    /*    ★実測＝ModelTables は はじめ 0／AddConnection は COM から 呼べない／
     *    データ分析は Microsoft の クラウド★ */
    三次元モデル:   function () { return 呼ぶ('三次元の窓を開く'); },
    Webから:        function () { return 呼ぶ('Webから読む'); },
    データモデル:   function () { return 呼ぶ('データモデルを開く'); },
    データ分析:     function () { return 呼ぶ('データ分析'); },

    /* ── ページ レイアウト ── */
    /* ★実Excel実測＝縦／A4／余白 上下19.1mm・左右17.8mm／枠線も 見出しも 刷らない★ */
    印刷:           function () { return 呼ぶ('printSheet'); },
    ページ設定:     function () { return 呼ぶ('ページ設定を開く'); },
    印刷の向き:     function () { return 呼ぶ('向きを切り替える'); },
    枠線を刷る:     function () { return 呼ぶ('枠線を刷るか'); },
    見出しを刷る:   function () { return 呼ぶ('見出しを刷るか'); },
  };

  /* 字の大きさ＝画面の 入力欄を 1か所で 動かす */
  function _サイズ(delta) {
    var d = (typeof document !== 'undefined') ? document : null;
    if (!d) return false;
    var el = d.getElementById('fmt-font-size');
    if (!el) return false;
    var 今 = parseInt(el.value, 10) || 12;
    var 次 = Math.max(6, Math.min(72, 今 + delta));
    if (次 === 今) return true;
    el.value = String(次);
    if (typeof el.onblur === 'function') el.onblur({ target: el });
    else if (typeof W().applyFormat === 'function') W().applyFormat('fontSize', 次);
    return true;
  }

  return A;
}));
