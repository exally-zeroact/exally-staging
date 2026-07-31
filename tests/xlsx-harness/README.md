# 版対応 検証ハーネス（xlsx-harness）

Exallyのグリッドが出した答えが、**本物のExcelと同じか**を関数ごとに突き合わせ続ける仕掛け。
新機能ではない。ここは「うちの答えが正しいと言い切れる根拠」を置く場所。

## いちばん大事な決まり

- **未検証を緑と呼ばない。** 真値の無い版は「未検証」と出す。件数も別に数える。
- **LibreOffice は Excel の代用ではない。** 別の版の答えとして列を分ける。LO が一致しても「Excel一致」とは言わない。
- **計算は本番と同じ経路で行う。** 生の HyperFormula で測ると実測で 39件も答えが変わり、検証にならない。
- **既知の不一致は消さない。** 台帳(`known-diffs.json`)に載っている物は緑にするが、レポートには毎回全部並ぶ。

## 使い方

```bash
node tests/xlsx-harness/compare.mjs            # 検証してレポートを書く(新規の不一致があれば exit 1)
node tests/xlsx-harness/compare.mjs --self-test # わざと壊して、ちゃんと赤くなるかを4通り確認
node tests/xlsx-harness/roundtrip.test.mjs     # 数式入りxlsxの往復(SheetJS)
```

Windows + Excel のある機械でだけ動く物（CIでは走らない）:

```powershell
pwsh -File tools/golden-excel.ps1 -Init     # 真値(golden)を初回作成
pwsh -File tools/golden-excel.ps1           # 2回目以降=pending と差分を出す(★既存を上書きしない)
pwsh -File tools/roundtrip-excel.ps1        # 往復(3): 書き出したxlsxを実Excelで開いて再計算
pwsh -File tools/verify-workbook-excel.ps1  # 同上を全221ケースで
```

## ファイル

| ファイル | 役目 |
|---|---|
| `cases/_inputs.json` | 全ケース共通の入力セル（★型を明示）と、入力の解釈を見るプローブ |
| `cases/*.json` | 検証ケース本体（52関数 + `&` 演算子・223ケース） |
| `golden/excel-365-*.json` | **真値**。実Excelで作った固定資産 |
| `golden/RECIPE.md` | どの環境でどう作ったか＝再現手順と記録 |
| `known-diffs.json` | 既知の不一致の台帳（区分A/B/C・区分Aは期限必須） |
| `route-snapshot.json` | 経路の錠。生HFに落ちたら赤くなる |
| `run-exally.mjs` | book.html を jsdom に載せて**本番経路**で計算する |
| `compare.mjs` | 突合・判定・レポート・赤/緑 |
| `report.md` | 生成物（人が読める表） |
| `xlsx-io.js` | 数式入りxlsxの書き出し/読み戻しの唯一の口 |
| `roundtrip.test.mjs` | 往復(1)(2) |
| `build-libre-input.mjs` / `collect-libre.mjs` | LibreOffice 用（別ジョブ） |

## xlsx-io.js の置き場所について

今はわざと `tests/xlsx-harness/` の下に置いている。`lib/` に置くと `scripts/stamp-build.mjs` が
`lib/` 配下の全 .js を内容ハッシュに含めるため、**どのHTMLからも読まれていないのに全HTMLの `?v=` が動く**。
誰も使わないコードを配信物に混ぜないための措置。
**P2でグリッド(book.html)が実際にこれを呼ぶ時に `lib/` へ移す。その時が初めての配信。**
`book.html` の `saveXlsx()` は今もスタブのまま（今回は触っていない）。

## ★新しい関数と `_xlfn.`（実測でわかった一番大事な事）

XLOOKUP / IFS / TEXTJOIN / SORT / UNIQUE / FILTER のような新しい関数を、そのままの名前で xlsx に書くと
**Excelはそのファイルを開けない**（1本混ざっただけでブックごと開けない）。xlsx の中では `_xlfn.` を付けた
名前で保存する決まりのため。SheetJS は自動では付けない。

`xlsx-io.js` が書き出し時に付け、読み戻し時に外す。対象の一覧は
「`_xlfn.` を付けたら実Excel(365 16.0.20228)が開いて正しく計算した」ことを**1関数ずつ確かめた結果**（37関数）。

`LET` / `LAMBDA` は関数名だけでなく引数名にも `_xlpm.` が要る
（`_xlfn.LET(_xlpm.x,2,_xlpm.x*3)` = 6 を実測）。引数名の付け替えは未対応なので、
**黙って壊れたファイルを作らないよう書き出しを止める**（例外を投げる）。

## jsdom で測った値が実ブラウザと同じか

ハーネスは jsdom の上で book.html を動かす。jsdom が実ブラウザと違う答えを出したら
検証そのものが嘘になるので、実ブラウザでも同じ手順を回して突き合わせた。

- ローカルに配信して（例: `python -m http.server 8791`）Chrome で `book.html` を開く
- 開発者コンソールで `cases/*.json` を fetch → `initFormulaEngine` → `setCellFormula` で全ケースを回す
- 出た値を `report.md` の「Exally(本番経路)」列と突き合わせる

2026-07-31 実施: **221/221 完全一致**（実Chrome と jsdom）。

## 不一致が出たらどう直すか

`known-diffs.json` に区分を書く。

- **A = うちで潰す** … `exally-formula.js` の `_jsComputeFormula` の対象表(`_jsSet`)に足して実装する。
  **入口はここ一箇所だけ。別経路を増やさない。** ★期限(`due`)必須・期限切れは赤。
- **B = HyperFormula 側の更新待ち** … 版と課題を書いて据え置く。
- **C = 仕様差として明示的に放置** … 理由を書いて表に出し続ける。
