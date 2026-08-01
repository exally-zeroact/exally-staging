# golden（真値）の作り方と記録

真値は **実Excel** で作る。ここに書いてあるのは「どの環境で・どうやって作ったか」。
環境が変われば答えも変わりうるので、作り直す時は必ずここを更新する。

## 1. 真値を作った環境（2026-07-31）

| 項目 | 値 |
|---|---|
| 製品 | Microsoft 365（`O365HomePremRetail`） |
| バージョン | **16.0.20228.20124** |
| プラットフォーム | x64 |
| 更新チャネル | `http://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60`（Current Channel） |
| UI言語 / インストール言語 | 1041 / 1041（日本語） |
| 地域（xlCountrySetting） | 81 |
| 小数点 / 桁区切り / リスト区切り | `.` / `,` / `,` |
| 日付システム | **1900系**（`Workbook.Date1904 = False`） |
| 実行 | `pwsh -File tools/golden-excel.ps1 -Init` |

> ★Current Channel は勝手に上がる。**真値が動く前提**でバージョンを毎回記録する。

### 決まりごと（守らないと真値がズレる）

- **式は `.Formula2`（US-English構文）で入れる。`.FormulaLocal` は使わない**（環境で式が変わるため）。
  - ★2026-08-01 修正: それまで `.Formula`（旧来インターフェース）を使っていた。非配列の式では違いが出ないが、
    **配列を返す式では「暗黙の交差」が働き、置いた行によって答えが変わる**（実測: `=SUM(C1:C6*E1:E6)` が
    row1で100 / row2で0 …。`.Formula2` なら 97100 で位置に依存しない＝人がExcelの画面で入力した時と同じ）。
  - 切り替え時に **既存245ケースの真値は1件も変わらなかった**ことを DIFF で確認済み。
  - 古いExcelには `Formula2` が無いので、その場合は `.Formula` に自動で落とす。
- **文字列セルは `NumberFormat='@'` を先に当ててから書く。**
  でないと `0007` / `1,234` / `2026-07-31` が数値・日付へ化けて、比べている物が変わる（実際に踏んだ）。
- **1900年の閏年バグ**（1900/2/29 が serial 60 として存在する）は Excel の仕様として受け入れる。
  ケースは **serial 61（1900/3/1）以降のみ**を使い、60以前は区分Cで別枠管理する。
- `TODAY` / `NOW` は golden に入れない（毎回変わるため）。実行時点との一致だけを見る。
- **スピルする式（`spill: true`）は1本ずつ別の列に置く。** T列に並べると隣の行へ溢れて次の式を潰し、
  `#SPILL!` だらけになる。Excelが何セルを占めたかは `spillCells` に記録する（うちのグリッドは再現できない＝区分C）。

## 2. エラー値の数値コード → 表示

Excel の COM は `#DIV/0!` などを大きな負の数で返す。対応表は
**Microsoft公式の enum 値**と**この環境での実測値**の両方で確かめてある。

出典: Microsoft Learn「XlCVError enumeration (Excel)」
<https://learn.microsoft.com/en-us/office/vba/api/excel.xlcverror>

| 表示 | 公式 enum 値 | COM Value2（実測） | 確認状態 |
|---|---|---|---|
| `#NULL!`  | xlErrNull 2000 | -2146826288 | 公式＋実測一致 |
| `#DIV/0!` | xlErrDiv0 2007 | -2146826281 | 公式＋実測一致 |
| `#VALUE!` | xlErrValue 2015 | -2146826273 | 公式＋実測一致 |
| `#REF!`   | xlErrRef 2023 | -2146826265 | 公式＋実測一致 |
| `#NAME?`  | xlErrName 2029 | -2146826259 | 公式＋実測一致 |
| `#NUM!`   | xlErrNum 2036 | -2146826252 | 公式＋実測一致 |
| `#N/A`    | xlErrNA 2042 | -2146826246 | 公式＋実測一致 |
| `#SPILL!` | xlErrSpill 2045 | （予測 -2146826243） | **公式のみ・未実測** |
| `#CALC!`  | **公式ページに記載なし** | -2146826238（2050 相当） | **実測のみ・出典なし** |

導出規則も実測で確認: **COM Value2 = -2146828288 + XlCVError値**（8件中7件で一致）。
ただし表は導出ではなく**実測値をそのまま**持つ（規則が破れた時に気付けるように）。

## 3. 作り直す時（★勝手に上書きしない）

```powershell
pwsh -File tools/golden-excel.ps1
```

既存の golden があると、上書きせずに

- `golden/pending-<日付>.json` … 今回のExcelが出した値
- `golden/DIFF-<日付>.md` … 既存との差分

を書く。**差分が出た＝Excel側が変わったという重大な情報。報告してから差し替える。**
差分ゼロなら pending は捨てる。

## 4. 他の版（未検証を埋めたい時）

| 版 | 状態 | 埋め方 |
|---|---|---|
| Excel 365 (16.0.20228) | **真値** | 済 |
| Excel 2016 / 2019 / Mac | **未検証** | その環境で `tools/golden-excel.ps1 -Init` を1回走らせて、出た JSON を `golden/` に置く。ファイル名の版が違うので既存とは別ファイルになる |
| LibreOffice | **未検証** | Linux で `bash tools/golden-libre.sh`（CIの別ジョブ・週次）。★Excelの真値ではない。参考列 |

`golden-excel.ps1` は **pwsh（PowerShell 7）** で動かす。Windows PowerShell 5.1 でも動くが、
その場合はこのスクリプトを BOM 付きで保存すること（日本語コメントが化けるため）。

## 5. 同梱している SheetJS

| 項目 | 値 |
|---|---|
| 版 | **SheetJS CE 0.20.3** |
| ライセンス | **Apache-2.0**（package.json の `license` を実測） |
| 入手元 | <https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js> |
| バイト数 | 951,904 |
| **sha256** | `cc015130aa8521e7f088f88898eba949ccdcbfb38df0bd129b44b7273c3a6f41` |
| 置き場所 | `lib/xlsx.full.min.js`（共有ライブラリの置き場。`kyuyo/` からは `../lib/...` で参照＝スタンプ対象の形） |

以前は `cdn.sheetjs.com/xlsx-latest/...` を読んでいた＝**版が勝手に動く**指定だった。
版検証をやるのにライブラリが浮動では意味が薄いので固定した。

差し替える時は sha256 も必ず更新する（版番号とバイト数だけでは中身の入れ替わりに気付けない）。

```bash
sha256sum lib/xlsx.full.min.js
```

## 6. `_xlfn.` の一覧をどう作ったか

「その名前のまま xlsx に書くと Excel がファイルを開けない」関数を、
**1関数ずつ実Excelで開いて確かめた**（プレフィックス無し → `_xlfn.` → `_xlfn._xlws.` の順に試し、
ファイルが開いて `#NAME?` にならない形を採用）。結果は 38関数中 **37関数が `_xlfn.`**。
`SORT` / `FILTER` は `_xlfn._xlws.` でも通るが、`_xlfn.` で通ることを実測したのでそちらに統一した。
`LET` / `LAMBDA` だけはどの接頭辞でも開けず、引数名に `_xlpm.` が要ることを実測で確認した
（`_xlfn.LET(_xlpm.x,2,_xlpm.x*3)` = 6 / `_xlfn.LAMBDA(_xlpm.y,_xlpm.y*2)(4)` = 8）。
