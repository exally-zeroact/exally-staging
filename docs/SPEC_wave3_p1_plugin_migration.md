# 【設計だけ・実装GO待ち】第3波 P1：11関数を HF プラグインへ移す

作成 2026-08-01 / 対象 `exally-staging` / **コードは1行も書いていない**

## 0. なぜやるか（1行）

`_jsSet` に残っている関数は**式の一番外側でしか効かない**。`=ROUND(LOOKUP(...),0)` のように入れ子で使われると
HyperFormula の答え（多くは `#NAME?`）に戻る＝**表示だけ正しくて参照先が間違う**。P1の11関数はそれが実務で起きる。

現状：入れ子で答えが変わる式 **38本**（`nesting-audit.mjs --probe`）。P1完了で **27本**まで下がる見込み。

## 1. 進め方（毎回この順）

1. **真値ケースを先に足す**（下の表の式を `cases/*.json` へ）
2. `pwsh -File tools/golden-excel.ps1` → DIFF が「**新規ケースのみ・既存値の変化ゼロ**」を確認してから昇格
3. その時点では**赤**（実装前だから当然）。ここで初めて実装する
4. 実装 → `compare.mjs` 緑 → `nesting-audit --check` の件数が下がったら `--update-baseline` で締め直す
5. 2層検証（全ケース実データ／実UIで全ボタン）＋ 実Chrome突合 ＋ 実Excel突合 → CI全通し

**★1関数ずつではなく11関数まとめて1コミット**（段階導入禁止）。ただし上の1〜3は関数ごとに回す。

## 2. 移す先の判断（ここが設計の肝）

| 関数 | 移す先 | 理由 |
|---|---|---|
| CONCAT / FIXED / DATEVALUE / NUMBERVALUE / ASC / JIS / TEXTBEFORE / TEXTAFTER | **プラグイン** | 純粋関数（入力→出力だけ）。入れ子で使われる |
| LOOKUP / XMATCH | **プラグイン** | 範囲を値として受け取れば足りる（参照を返さない） |
| **INDIRECT** | **★要スパイク。プラグインで足りない可能性が高い** | 下の §4 |

## 3. 関数ごとの仕様と「何を測れば合っていると言えるか」

真値は毎回 **実Excel（16.0.20228 / 日本語1041）** から取る。ケースIDは `<関数>_<観点>`。

### CONCAT
- Excel: 範囲も引数に取れる。空セルは飛ばさず""として連結。数値は既定書式で文字列化。
- 測る観点 → ケース案
  - `CONCAT_range` : `=CONCAT(D1:D3)` … 範囲を渡せるか（現行の主用途）
  - `CONCAT_mixed` : `=CONCAT(B1,"-",A1)` … 文字と数値の混在
  - `CONCAT_blank` : `="["&CONCAT(G1:G3)&"]"` … 空セルの扱い（CONCATENATEと違い範囲可）
  - `CONCAT_nested` : `=LEN(CONCAT(D1:D6))` … ★入れ子
  - `CONCAT_number_fmt` : `=CONCAT(A4)` … 0.1 が "0.1" になるか（書式が付かないこと）

### LOOKUP
- Excel: ベクトル形式。**検索範囲は昇順前提**、見つからなければ「超えない最大」。全部より小さければ `#N/A`。
- 観点 → ケース案
  - `LOOKUP_exact` : `=LOOKUP(20,C1:C6,E1:E6)`
  - `LOOKUP_between` : `=LOOKUP(30,C1:C6,E1:E6)` … 中間値＝超えない最大を返すか（ここが一番間違えやすい）
  - `LOOKUP_below_all` : `=IFERROR(LOOKUP(0,C1:C6,E1:E6),"NA")` … 全部より小さい
  - `LOOKUP_above_all` : `=LOOKUP(999,C1:C6,E1:E6)` … 全部より大きい＝最後
  - `LOOKUP_text` : `=LOOKUP("B",D1:D6,E1:E6)` … 文字列の比較順
  - `LOOKUP_nested` : `=ROUND(LOOKUP(30,C1:C6,E1:E6),0)` … ★入れ子

### XMATCH
- Excel: `XMATCH(値, 配列, [一致モード], [検索モード])`。既定は完全一致（MATCHの既定と違う）。
- 観点 → ケース案
  - `XMATCH_exact` : `=XMATCH(20,C1:C6)` … 既定が完全一致であること
  - `XMATCH_miss` : `=IFERROR(XMATCH(30,C1:C6),"NA")` … 既定で見つからなければ #N/A
  - `XMATCH_next_smaller` : `=XMATCH(30,C1:C6,-1)` … 一致モード -1
  - `XMATCH_next_larger` : `=XMATCH(30,C1:C6,1)` … 一致モード 1
  - `XMATCH_wildcard` : `=XMATCH("りん*",B1:B8,2)` … 一致モード 2 でワイルドカード
  - `XMATCH_reverse` : `=XMATCH("A",D1:D6,0,-1)` … 検索モード -1（後ろから）
  - `XMATCH_nested` : `=INDEX(E1:E6,XMATCH(20,C1:C6))` … ★入れ子

### DATEVALUE
- Excel: 文字列 → 日付シリアル。**ロケール依存**（`2026/7/31`・`2026-07-31`・`R8.7.31` 等）。時刻付きは小数を切り捨て。
- 観点 → ケース案
  - `DATEVALUE_slash` : `=DATEVALUE("2026/7/31")`
  - `DATEVALUE_hyphen` : `=DATEVALUE("2026-07-31")`
  - `DATEVALUE_jp` : `=IFERROR(DATEVALUE("2026年7月31日"),"NA")` … 和暦・日本語表記でどうなるか**実Excelに聞く**
  - `DATEVALUE_bad` : `=IFERROR(DATEVALUE("あ"),"NA")`
  - `DATEVALUE_nested` : `=YEAR(DATEVALUE("2026-07-31"))` … ★入れ子
- ★注意：**推測で実装しない**。どの表記を受けるかは実Excelの答えを golden にしてから決める。

### NUMBERVALUE
- Excel: `NUMBERVALUE(文字列, [小数点], [桁区切り])`。区切りを明示できるのが VALUE との違い。
- 観点 → ケース案
  - `NUMBERVALUE_plain` : `=NUMBERVALUE("1.5")`
  - `NUMBERVALUE_sep` : `=NUMBERVALUE("1,234.5")` … 既定の区切り
  - `NUMBERVALUE_custom` : `=NUMBERVALUE("1.234,5",",",".")` … 欧州式（引数で区切りを指定）
  - `NUMBERVALUE_bad` : `=IFERROR(NUMBERVALUE("あ"),"NA")`
  - `NUMBERVALUE_nested` : `=SUM(NUMBERVALUE("1,234"),1)` … ★入れ子

### FIXED
- Excel: `FIXED(数値, [桁数], [桁区切りを付けない])`。**四捨五入して文字列**を返す。
- 観点 → ケース案
  - `FIXED_default` : `=FIXED(1234.567)` … 既定2桁＋桁区切り
  - `FIXED_digits0` : `=FIXED(1234.5,0)` … 丸め方向（0から遠い方）
  - `FIXED_nocomma` : `=FIXED(1234.567,2,TRUE)` … 桁区切りなし
  - `FIXED_negative_digits` : `=FIXED(1234.5,-2)` … 負の桁数
  - `FIXED_is_text` : `=LEN(FIXED(1234.5,0))` … ★文字列で返ること＋入れ子

### ASC / JIS
- Excel: 全角→半角 / 半角→全角。**かな・記号・スペースも対象**。日本語環境の実務で最頻出。
- 観点 → ケース案（入力セルに全角文字を足す必要あり → `_inputs.json` に `B9="ＡＢＣ１２３"`, `B10="ｱｲｳ"` を追加）
  - `ASC_alnum` : `=ASC(B9)` … 全角英数 → 半角
  - `ASC_kana` : `=ASC("アイウ")` … 全角カナ → 半角カナ
  - `ASC_mixed` : `=ASC("Ａ亜１")` … 漢字は変換されないこと
  - `JIS_alnum` : `=JIS("ABC123")` … 半角 → 全角
  - `JIS_kana` : `=JIS(B10)` … 半角カナ → 全角カナ（濁点の合成に注意）
  - `ASC_nested` : `=LEN(ASC(B9))` … ★入れ子（長さが変わるか＝変換が効いたかが1数字で分かる）
- ★注意：濁点・半濁点（`ｶﾞ`→`ガ`）の合成は実装がずれやすい。**必ず実Excelの答えを真値にする**。

### TEXTBEFORE / TEXTAFTER
- Excel: `TEXTBEFORE(文字列, 区切り, [出現回数], [大小区別], [末尾一致], [見つからない時])`。
- 観点 → ケース案
  - `TEXTBEFORE_first` : `=TEXTBEFORE("007-1234","-")`
  - `TEXTAFTER_first` : `=TEXTAFTER("007-1234","-")`
  - `TEXTBEFORE_nth` : `=TEXTBEFORE("a-b-c","-",2)` … 2番目の区切り
  - `TEXTAFTER_negative` : `=TEXTAFTER("a-b-c","-",-1)` … 後ろから
  - `TEXTBEFORE_missing` : `=IFERROR(TEXTBEFORE("abc","-"),"NA")` … 区切りが無い
  - `TEXTBEFORE_ifmissing` : `=TEXTBEFORE("abc","-",1,0,0,"なし")` … 見つからない時の既定値
  - `TEXTAFTER_nested` : `=LEN(TEXTAFTER("007-1234","-"))` … ★入れ子

## 4. ★INDIRECT は別扱い（ここだけ設計を確定させない）

INDIRECT は**値ではなく参照を返す**関数。`=SUM(INDIRECT("E1:E6"))` は「範囲」を渡せないと成立しない。
HyperFormula のプラグインは基本「値を返す」形なので、そのまま実装すると
`=INDIRECT("E1")` は動いても `=SUM(INDIRECT("E1:E6"))` が動かない、という半端な物になる恐れがある。

**やること（実装GOの前に1つだけ）**：小さなスパイクで
`=SUM(INDIRECT("E1:E6"))` / `=INDIRECT("E1")` / `=INDIRECT(B5)` の3本が
プラグインで成立するかを実測する。成立しないなら

- (a) `_jsSet` に残し、「一番外側でしか効かない」ことを台帳に明記して期限を切る
- (b) グリッド側（book.html）で式を書き換える別方式にする

のどちらかを選んで**その時に報告する**。★推測で実装に入らない。

## 5. 入力セルの追加（真値の作り直しが要る）

ASC / JIS のために `cases/_inputs.json` に足す：

| セル | 値 | 用途 |
|---|---|---|
| B9 | `ＡＢＣ１２３` | 全角英数 |
| B10 | `ｱｲｳ` | 半角カナ |
| B11 | `ｶﾞｷﾞ` | 半角カナ＋濁点（合成の確認） |

**入力を足すと既存ケースの真値が変わらないことを DIFF で確認する**（B列の範囲を使う既存ケースに影響が出ないか。
`LEN_jp` 等は個別セル参照なので影響しないはずだが、**確認してから**進める）。

## 6. 完了の判定（数字で出す）

- `compare.mjs`：一致 **243 → 約280**（新規ケース約37件がすべて一致）／新規不一致 0
- `nesting-audit --probe`：**38 → 27**（P1の11関数が消える）→ baseline を締め直す
- `--self-test` 6通り／`roundtrip` 14/14／`hub-ui` 65/65／stamp ／CI全通し
- 実Chromeで全ケース突合＝jsdomと一致、実Excelで書き出しブック突合＝全件一致

## 7. やらないこと

- コードは書かない（この設計の承認後）
- INDIRECT の実装方式を今決めない（§4のスパイク結果で決める）
- P2 / P3 の関数には手を付けない
